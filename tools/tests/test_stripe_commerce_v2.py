from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pytest


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import release_product  # noqa: E402
from tools.commerce import stripe_commerce_v2 as commerce  # noqa: E402
from tools.store.stripe_release_core import (  # noqa: E402
    DESKTOP_DOWNLOAD_REDIRECT_URL,
    app_delivery_config,
    base_payload,
    build_app_release,
    sha256_payload,
)


class Page(list):
    def auto_paging_iter(self):
        return iter(self)


class ProductAPI:
    def __init__(self, owner: "FakeStripe") -> None:
        self.owner = owner

    def list(self, **_: Any) -> Page:
        self.owner.api_calls.append("Product.list")
        return Page(self.owner.products.values())

    def create(self, idempotency_key: str, **payload: Any) -> dict[str, Any]:
        self.owner.api_calls.append("Product.create")
        self.owner.idempotency_keys.append(idempotency_key)
        value = {**payload, "id": f"prod_{len(self.owner.products) + 1}", "livemode": True}
        self.owner.products[value["id"]] = value
        return value

    def retrieve(self, object_id: str) -> dict[str, Any]:
        self.owner.api_calls.append("Product.retrieve")
        value = dict(self.owner.products[object_id])
        if self.owner.bad_product_readback:
            value["name"] = "wrong read-back"
        return value


class PriceAPI:
    def __init__(self, owner: "FakeStripe") -> None:
        self.owner = owner

    def list(self, product: str, **_: Any) -> Page:
        self.owner.api_calls.append("Price.list")
        return Page(value for value in self.owner.prices.values() if value["product"] == product and value["active"])

    def create(self, idempotency_key: str, **payload: Any) -> dict[str, Any]:
        self.owner.api_calls.append("Price.create")
        self.owner.idempotency_keys.append(idempotency_key)
        value = {
            **payload,
            "id": f"price_{len(self.owner.prices) + 1}",
            "active": True,
            "livemode": True,
            "type": "one_time",
            "recurring": None,
        }
        self.owner.prices[value["id"]] = value
        return value

    def retrieve(self, object_id: str) -> dict[str, Any]:
        self.owner.api_calls.append("Price.retrieve")
        return dict(self.owner.prices[object_id])


class PaymentLinkAPI:
    def __init__(self, owner: "FakeStripe") -> None:
        self.owner = owner

    def list(self, **_: Any) -> Page:
        self.owner.api_calls.append("PaymentLink.list")
        return Page(value for value in self.owner.links.values() if value["active"])

    def create(self, idempotency_key: str, **payload: Any) -> dict[str, Any]:
        self.owner.api_calls.append("PaymentLink.create")
        self.owner.idempotency_keys.append(idempotency_key)
        link_id = f"plink_{len(self.owner.links) + 1}"
        lines = [dict(line) for line in payload.pop("line_items")]
        value = {
            **payload,
            "id": link_id,
            "url": f"https://buy.stripe.com/{link_id}",
            "active": True,
            "livemode": True,
        }
        self.owner.links[link_id] = value
        self.owner.link_lines[link_id] = lines
        return value

    def retrieve(self, object_id: str) -> dict[str, Any]:
        self.owner.api_calls.append("PaymentLink.retrieve")
        return dict(self.owner.links[object_id])

    def list_line_items(self, object_id: str, **_: Any) -> Page:
        self.owner.api_calls.append("PaymentLink.list_line_items")
        return Page(self.owner.link_lines[object_id])


class FakeStripe:
    def __init__(self) -> None:
        self.api_key = ""
        self.api_calls: list[str] = []
        self.idempotency_keys: list[str] = []
        self.products: dict[str, dict[str, Any]] = {}
        self.prices: dict[str, dict[str, Any]] = {}
        self.links: dict[str, dict[str, Any]] = {}
        self.link_lines: dict[str, list[dict[str, Any]]] = {}
        self.bad_product_readback = False
        self.Product = ProductAPI(self)
        self.Price = PriceAPI(self)
        self.PaymentLink = PaymentLinkAPI(self)


def payload(product_id: str = "example_app") -> dict[str, Any]:
    product_payload = {
        "name": "Example App",
        "description": "Small useful app.",
        "active": True,
        "tax_code": "txcd_10202003",
        "metadata": {"dake_item_id": product_id, "source_original": f"01_apps/{product_id}/ORIGINAL.md"},
    }
    price_payload = {
        "currency": "jpy",
        "unit_amount": 300,
        "product": commerce.PRODUCT_PLACEHOLDER,
        "metadata": {"dake_item_id": product_id},
    }
    link_payload = {
        "line_items": [{"price": commerce.PRICE_PLACEHOLDER, "quantity": 1}],
        "metadata": {"dake_item_id": product_id},
        "payment_intent_data": {"metadata": {"dake_item_id": product_id}},
    }
    return {
        "mode": "dry-run",
        "product_id": product_id,
        "product_type": "app",
        "source_original": f"01_apps/{product_id}/ORIGINAL.md",
        "title": "Example App",
        "price": 300,
        "currency": "jpy",
        "product_payload": product_payload,
        "price_payload": price_payload,
        "payment_link_payload": link_payload,
        "product_payload_sha256": sha256_payload(product_payload),
        "price_payload_sha256": sha256_payload(price_payload),
        "payment_link_payload_sha256": sha256_payload(link_payload),
        "product_idempotency_key": "commerce-product-v2-example",
        "price_idempotency_key": "commerce-price-v2-example",
        "payment_link_idempotency_key": "commerce-link-v2-example",
        "ready_for_live_execution": "yes",
        "secret_read": "no",
        "live_api_called": "no",
        "errors": [],
    }


def test_desktop_private_delivery_adds_canonical_redirect() -> None:
    original = """
- purchase_delivery_method: stripe_session_private_r2
- purchase_delivery_ready: yes
- delivery_object_key: desktop/DAKE_Mail_Draft/current.zip
"""
    delivery = app_delivery_config(original, "DAKE_Mail_Draft")
    assert delivery == {
        "method": "stripe_session_private_r2",
        "ready": "yes",
        "object_key": "desktop/DAKE_Mail_Draft/current.zip",
        "after_completion_url": DESKTOP_DOWNLOAD_REDIRECT_URL,
        "errors": [],
    }

    _, _, payment_link = base_payload(
        product_id="DAKE_Mail_Draft",
        product_type="app",
        title="Dakeメール下書き",
        description="メール下書きを作るDesktop DAKE。",
        price=300,
        currency="jpy",
        source_original="01_apps/DAKE_Mail_Draft/ORIGINAL.md",
        booth_url="",
        github_release_url="",
        tax_code_candidate="txcd_10202003",
        purchase_delivery_method=delivery["method"],
        after_completion_url=delivery["after_completion_url"],
    )
    assert payment_link["after_completion"] == {
        "type": "redirect",
        "redirect": {"url": DESKTOP_DOWNLOAD_REDIRECT_URL},
    }
    assert payment_link["metadata"]["dake_item_id"] == "DAKE_Mail_Draft"
    assert "delivery_object_key" not in json.dumps(payment_link)


@pytest.mark.parametrize(
    ("ready", "object_key"),
    [
        ("no", "desktop/DAKE_Mail_Draft/current.zip"),
        ("yes", "desktop/other/current.zip"),
        ("yes", "../private.zip"),
    ],
)
def test_desktop_private_delivery_stops_when_not_exact(ready: str, object_key: str) -> None:
    original = f"""
- purchase_delivery_method: stripe_session_private_r2
- purchase_delivery_ready: {ready}
- delivery_object_key: {object_key}
"""
    delivery = app_delivery_config(original, "DAKE_Mail_Draft")
    assert delivery["ready"] == "no"
    assert delivery["after_completion_url"] is None
    assert delivery["errors"]


def test_app_without_private_delivery_keeps_existing_payment_link_payload() -> None:
    delivery = app_delivery_config("# ORIGINAL\n", "example_app")
    assert delivery["ready"] == "not_applicable"
    assert delivery["after_completion_url"] is None


def test_app_release_payload_uses_original_delivery_fields() -> None:
    product_id = "DAKE_Mail_Draft"
    result = build_app_release(
        {
            "product_id": product_id,
            "source_original": f"01_apps/{product_id}/ORIGINAL.md",
            "original_text": """
- purchase_delivery_method: stripe_session_private_r2
- purchase_delivery_ready: yes
- delivery_object_key: desktop/DAKE_Mail_Draft/current.zip
""",
            "item": {
                "title": "Dakeメール下書き",
                "description": "メール下書きを作るDesktop DAKE。",
                "price": 300,
                "currency": "JPY",
                "payment_status": "booth_only",
                "stripe_payment_link": "",
            },
        }
    )
    assert result["ready_for_live_execution"] == "yes"
    assert result["purchase_delivery_method"] == "stripe_session_private_r2"
    assert result["delivery_object_key"] == "desktop/DAKE_Mail_Draft/current.zip"
    assert result["payment_link_payload"]["after_completion"]["redirect"]["url"] == DESKTOP_DOWNLOAD_REDIRECT_URL


def prepare_root(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, data: dict[str, Any]) -> Path:
    monkeypatch.setattr(commerce, "ROOT", tmp_path)
    monkeypatch.setattr(commerce, "ARTIFACT_ROOT", tmp_path / "tools" / "reports" / "release_artifacts")
    original = tmp_path / data["source_original"]
    original.parent.mkdir(parents=True, exist_ok=True)
    original.write_text(
        "# ORIGINAL\n\n- payment_status: booth_only\n- stripe_payment_link: 未設定\n",
        encoding="utf-8",
    )
    return original


def add_product(fake: FakeStripe, data: dict[str, Any], object_id: str = "prod_existing") -> dict[str, Any]:
    value = {**data["product_payload"], "id": object_id, "livemode": True}
    fake.products[object_id] = value
    return value


def add_price(fake: FakeStripe, data: dict[str, Any], product_id: str, object_id: str = "price_existing") -> dict[str, Any]:
    value = {
        **data["price_payload"],
        "product": product_id,
        "id": object_id,
        "active": True,
        "livemode": True,
        "type": "one_time",
        "recurring": None,
    }
    fake.prices[object_id] = value
    return value


def add_link(fake: FakeStripe, data: dict[str, Any], price_id: str, object_id: str = "plink_existing") -> dict[str, Any]:
    value = {
        "id": object_id,
        "url": f"https://buy.stripe.com/{object_id}",
        "active": True,
        "livemode": True,
        "metadata": dict(data["payment_link_payload"]["metadata"]),
        "payment_intent_data": dict(data["payment_link_payload"]["payment_intent_data"]),
    }
    for key in ("custom_text", "after_completion"):
        if key in data["payment_link_payload"]:
            value[key] = data["payment_link_payload"][key]
    fake.links[object_id] = value
    fake.link_lines[object_id] = [{"price": price_id, "quantity": 1}]
    return value


def run_live(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, fake: FakeStripe, data: dict[str, Any] | None = None):
    data = data or payload()
    original = prepare_root(monkeypatch, tmp_path, data)
    result = commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    return result, original


def test_dry_run_reads_no_secret_and_calls_no_api(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    prepare_root(monkeypatch, tmp_path, data)
    monkeypatch.setattr(commerce, "build_release_payload", lambda _: data)
    monkeypatch.setattr(commerce, "validate_live_secret_key", lambda: (_ for _ in ()).throw(AssertionError("secret read")))
    monkeypatch.setattr(commerce, "import_stripe_module", lambda: (_ for _ in ()).throw(AssertionError("API import")))
    _, plan = commerce.build_dry_run_plan(data["product_id"])
    assert plan["secret_read"] == "no"
    assert plan["live_api_called"] == "no"
    assert fake.api_calls == []


def test_test_key_is_rejected_before_api(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    prepare_root(monkeypatch, tmp_path, data)
    with pytest.raises(commerce.SafetyStop, match="test mode"):
        commerce.execute_live(data, stripe_client=fake, secret_key="sk_test_fixture_1234567890")
    assert fake.api_calls == []


def test_product_none_creates_all_and_updates_only_on_success(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    fake = FakeStripe()
    result, original = run_live(monkeypatch, tmp_path, fake)
    assert result["actions"] == {"product": "created", "price": "created", "payment_link": "created", "original": "updated"}
    text = original.read_text(encoding="utf-8")
    assert "payment_status: stripe_ready" in text
    assert result["payment_link_url"] in text
    assert len(fake.idempotency_keys) == 3


def test_exact_product_reused_and_partial_product_state_continues(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    add_product(fake, data)
    result, _ = run_live(monkeypatch, tmp_path, fake, data)
    assert result["actions"]["product"] == "reused"
    assert "Product.create" not in fake.api_calls
    assert result["actions"]["price"] == "created"


def test_multiple_products_stop(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    add_product(fake, data, "prod_1")
    add_product(fake, data, "prod_2")
    original = prepare_root(monkeypatch, tmp_path, data)
    before = original.read_text(encoding="utf-8")
    with pytest.raises(commerce.SafetyStop, match="ambiguous Stripe Products"):
        commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    assert original.read_text(encoding="utf-8") == before


def test_exact_price_reused_and_product_price_partial_state_continues(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    add_product(fake, data)
    add_price(fake, data, "prod_existing")
    result, _ = run_live(monkeypatch, tmp_path, fake, data)
    assert result["actions"]["price"] == "reused"
    assert "Price.create" not in fake.api_calls
    assert result["actions"]["payment_link"] == "created"


def test_nonmatching_price_is_not_mutated_and_new_price_is_created(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    add_product(fake, data)
    wrong = add_price(fake, data, "prod_existing")
    wrong["unit_amount"] = 999
    result, _ = run_live(monkeypatch, tmp_path, fake, data)
    assert result["actions"]["price"] == "created"
    assert fake.prices["price_existing"]["unit_amount"] == 999


def test_exact_payment_link_reused(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    add_product(fake, data)
    add_price(fake, data, "prod_existing")
    add_link(fake, data, "price_existing")
    result, _ = run_live(monkeypatch, tmp_path, fake, data)
    assert result["actions"]["payment_link"] == "reused"
    assert "PaymentLink.create" not in fake.api_calls


def test_mail_draft_private_delivery_reuses_product_and_price_and_replaces_old_link(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    product_id = "DAKE_Mail_Draft"
    data = build_app_release(
        {
            "product_id": product_id,
            "source_original": f"01_apps/{product_id}/ORIGINAL.md",
            "original_text": """
- purchase_delivery_method: stripe_session_private_r2
- purchase_delivery_ready: yes
- delivery_object_key: desktop/DAKE_Mail_Draft/current.zip
""",
            "item": {
                "title": "Dakeメール下書き",
                "description": "メール下書きを作るDesktop DAKE。",
                "price": 300,
                "currency": "JPY",
                "payment_status": "stripe_ready",
                "stripe_payment_link": "https://buy.stripe.com/old_mail_draft",
            },
        }
    )
    assert "purchase_delivery_method" not in data["product_payload"]["metadata"]
    assert "delivery_policy" not in data["product_payload"]["metadata"]

    fake = FakeStripe()
    add_product(fake, data, "prod_mail_draft")
    add_price(fake, data, "prod_mail_draft", "price_mail_draft_300")
    old_link = {
        "id": "plink_old_mail_draft",
        "url": "https://buy.stripe.com/old_mail_draft",
        "active": True,
        "livemode": True,
        "metadata": {
            "dake_item_id": product_id,
            "dake_type": "app",
            "source_original": f"01_apps/{product_id}/ORIGINAL.md",
        },
        "payment_intent_data": {"metadata": {"dake_item_id": product_id}},
    }
    fake.links[old_link["id"]] = old_link
    fake.link_lines[old_link["id"]] = [{"price": "price_mail_draft_300", "quantity": 1}]

    expected_link = {
        **data["payment_link_payload"],
        "line_items": [{"price": "price_mail_draft_300", "quantity": 1}],
    }
    candidate, candidate_action = commerce.select_payment_link(
        [(old_link, fake.link_lines[old_link["id"]])],
        "price_mail_draft_300",
        expected_link,
    )
    assert candidate is None
    assert candidate_action == "created"

    result, _ = run_live(monkeypatch, tmp_path, fake, data)
    assert result["actions"]["product"] == "reused"
    assert result["actions"]["price"] == "reused"
    assert result["actions"]["payment_link"] == "created"
    assert "Product.create" not in fake.api_calls
    assert "Price.create" not in fake.api_calls
    assert fake.links[old_link["id"]] == old_link

    new_link = fake.links[result["payment_link_id"]]
    assert new_link["metadata"]["purchase_delivery_method"] == "stripe_session_private_r2"
    assert new_link["payment_intent_data"]["metadata"]["delivery_policy"] == "stripe_session_private_r2"
    assert new_link["after_completion"] == {
        "type": "redirect",
        "redirect": {"url": DESKTOP_DOWNLOAD_REDIRECT_URL},
    }

    before_rerun = len(fake.links)
    second = commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    assert second["actions"]["payment_link"] == "reused"
    assert len(fake.links) == before_rerun


def test_multiple_exact_payment_links_stop(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    add_product(fake, data)
    add_price(fake, data, "prod_existing")
    add_link(fake, data, "price_existing", "plink_1")
    add_link(fake, data, "price_existing", "plink_2")
    original = prepare_root(monkeypatch, tmp_path, data)
    with pytest.raises(commerce.SafetyStop, match="ambiguous Stripe Payment Links"):
        commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    assert "stripe_ready" not in original.read_text(encoding="utf-8")


def test_readback_mismatch_does_not_update_original(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    fake.bad_product_readback = True
    original = prepare_root(monkeypatch, tmp_path, data)
    before = original.read_text(encoding="utf-8")
    with pytest.raises(commerce.SafetyStop, match="read-back mismatch"):
        commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    assert original.read_text(encoding="utf-8") == before


def test_rerun_reuses_everything_without_duplicates(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    result, original = run_live(monkeypatch, tmp_path, fake, data)
    counts = (len(fake.products), len(fake.prices), len(fake.links))
    second = commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    assert (len(fake.products), len(fake.prices), len(fake.links)) == counts
    assert second["actions"] == {"product": "reused", "price": "reused", "payment_link": "reused", "original": "unchanged"}
    assert result["payment_link_url"] in original.read_text(encoding="utf-8")


def test_no_checkout_or_production_review_is_needed_and_no_other_layer_runs(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    fake = FakeStripe()
    result, _ = run_live(monkeypatch, tmp_path, fake)
    assert result["commerce_status"] == "READY"
    artifact_names = {path.name for path in commerce.artifact_dir("example_app").iterdir()}
    assert artifact_names == {"commerce_state.json", "commerce_result.json"}
    assert not any(name in json.dumps(result) for name in ("SHIP", "BOOTH", "Cloudflare", "Buffer", "production_review", "checkout_review"))


def test_secret_like_error_is_redacted_from_state(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    data = payload()
    fake = FakeStripe()
    prepare_root(monkeypatch, tmp_path, data)
    secret = "sk_live_" + "1234567890abcdef"
    fake.Product.list = lambda **_: (_ for _ in ()).throw(RuntimeError(secret))  # type: ignore[method-assign]
    with pytest.raises(commerce.SafetyStop):
        commerce.execute_live(data, stripe_client=fake, secret_key="sk_live_fixture_1234567890")
    state_text = commerce.state_path(data["product_id"]).read_text(encoding="utf-8")
    assert secret not in state_text
    assert "redacted secret-like value" in state_text


def test_selectors_cover_create_reuse_and_ambiguity() -> None:
    data = payload()
    product, action = commerce.select_product([], data["product_id"], data["product_payload"])
    assert product is None and action == "created"
    exact_product = {**data["product_payload"], "id": "prod_x", "livemode": True}
    assert commerce.select_product([exact_product], data["product_id"], data["product_payload"])[1] == "reused"
    expected_price = {**data["price_payload"], "product": "prod_x"}
    price, action = commerce.select_price([], "prod_x", expected_price)
    assert price is None and action == "created"


def cli_plan() -> tuple[dict[str, Any], dict[str, Any]]:
    return payload(), {
        "title": "Example App",
        "amount": 300,
        "currency": "jpy",
        "ready": True,
        "errors": [],
    }


def test_apply_live_is_the_live_intent(monkeypatch: pytest.MonkeyPatch) -> None:
    called: list[str] = []
    monkeypatch.setattr(sys, "argv", ["release_product.py", "example_app", "--apply-live"])
    monkeypatch.setattr(release_product, "build_dry_run_plan", lambda _: cli_plan())
    monkeypatch.setattr(
        release_product,
        "execute_live",
        lambda _: called.append("live") or {
            "title": "Example App",
            "amount": 300,
            "currency": "jpy",
            "actions": {"product": "reused", "price": "reused", "payment_link": "reused", "original": "unchanged"},
        },
    )
    assert release_product.main() == 0
    assert called == ["live"]


def test_execute_live_legacy_flag_stops_before_plan_or_live(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    monkeypatch.setattr(sys, "argv", ["release_product.py", "example_app", "--execute-live"])
    monkeypatch.setattr(release_product, "build_dry_run_plan", lambda _: (_ for _ in ()).throw(AssertionError("plan started")))
    monkeypatch.setattr(release_product, "execute_live", lambda _: (_ for _ in ()).throw(AssertionError("live started")))
    assert release_product.main() == 1
    assert "--execute-live is legacy; use --apply-live" in capsys.readouterr().out


def test_cli_default_remains_dry_run(monkeypatch: pytest.MonkeyPatch) -> None:
    called: list[str] = []
    monkeypatch.setattr(sys, "argv", ["release_product.py", "example_app"])
    monkeypatch.setattr(release_product, "build_dry_run_plan", lambda _: cli_plan())
    monkeypatch.setattr(release_product, "execute_live", lambda _: called.append("live"))
    assert release_product.main() == 0
    assert called == []
