from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from typing import Any

from tools.store import sync_stripe_r2_links as mod


class FakeStripeClient:
    def __init__(
        self,
        *,
        links: list[dict[str, Any]],
        line_items_by_link: dict[str, list[dict[str, Any]]],
        products: list[dict[str, Any]],
        prices_by_product: dict[str, list[dict[str, Any]]] | None = None,
    ) -> None:
        self._links = links
        self._line_items = line_items_by_link
        self._products = products
        self._prices = prices_by_product or {}
        self.calls: list[tuple[str, Any]] = []
        self._seq = 0

    def list_payment_links(self, *, active: bool | None = None) -> list[dict[str, Any]]:
        if active is None:
            return list(self._links)
        return [l for l in self._links if l.get("active", True) is active]

    def list_payment_link_line_items(self, payment_link_id: str) -> list[dict[str, Any]]:
        return list(self._line_items.get(payment_link_id, []))

    def list_products(self) -> list[dict[str, Any]]:
        return list(self._products)

    def list_prices(self, product_id: str) -> list[dict[str, Any]]:
        return list(self._prices.get(product_id, []))

    def create_product(self, payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
        self._seq += 1
        prod = {"id": f"prod_new_{self._seq}", **payload}
        self.calls.append(("create_product", payload))
        return prod

    def create_price(self, payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
        self._seq += 1
        price = {"id": f"price_new_{self._seq}", **payload}
        self.calls.append(("create_price", payload))
        return price

    def create_payment_link(self, payload: dict[str, Any], idempotency_key: str) -> dict[str, Any]:
        self._seq += 1
        link = {
            "id": f"plink_new_{self._seq}",
            "url": f"https://buy.stripe.com/liveNew{self._seq}",
            "active": True,
            **payload,
        }
        self.calls.append(("create_payment_link", payload))
        return link

    def update_payment_link(self, payment_link_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        self.calls.append(("update_payment_link", payment_link_id, payload))
        return {"id": payment_link_id, **payload}


class SyncStripeR2LinksTest(unittest.TestCase):
    def test_flatten_form_params(self) -> None:
        pairs = mod.flatten_form_params({
            "line_items": [{"price": "price_123", "quantity": 1}],
            "after_completion": {"type": "redirect", "redirect": {"url": mod.R2_REDIRECT_URL}},
            "active": False,
        })
        self.assertIn(("line_items[0][price]", "price_123"), pairs)
        self.assertIn(("line_items[0][quantity]", "1"), pairs)
        self.assertIn(("after_completion[type]", "redirect"), pairs)
        self.assertIn(("after_completion[redirect][url]", mod.R2_REDIRECT_URL), pairs)
        self.assertIn(("active", "false"), pairs)

    def test_rejects_test_secret_key(self) -> None:
        with self.assertRaises(mod.StripeSyncError):
            mod.StripeClient("sk_test_1234567890abcdef")

    def test_secret_sanitization(self) -> None:
        raw = {"msg": "failed with sk_live_secret1234567890 and whsec_abcdefgh1234"}
        safe = mod.sanitize_obj(raw)
        self.assertNotIn("sk_live_secret", safe["msg"])
        self.assertNotIn("whsec_abcdefgh", safe["msg"])

    def test_live_reconcile_apply_handles_all_cases(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            repo_root = Path(tmpdir)
            for rel, content in [
                ("01_apps/AppSame/ORIGINAL.md", "- Stripe Payment Link: https://buy.stripe.com/same\n- payment_status: stripe_ready\n"),
                ("01_apps/AppNoop/ORIGINAL.md", "- Stripe Payment Link: https://buy.stripe.com/noop\n- payment_status: stripe_ready\n"),
                ("01_apps/AppPriceChange/ORIGINAL.md", "- Stripe Payment Link: https://buy.stripe.com/old300\n- payment_status: stripe_ready\n"),
                ("01_apps/AppNew/ORIGINAL.md", "- Stripe Payment Link: 未設定\n- payment_status: booth_only\n"),
            ]:
                p = repo_root / rel
                p.parent.mkdir(parents=True, exist_ok=True)
                p.write_text(content, encoding="utf-8")

            catalog = [
                {"id": "app_same", "title": "Same", "price": 500, "stripe_payment_link": "https://buy.stripe.com/same", "source_original": "01_apps/AppSame/ORIGINAL.md"},
                {"id": "app_noop", "title": "Noop", "price": 500, "stripe_payment_link": "https://buy.stripe.com/noop", "source_original": "01_apps/AppNoop/ORIGINAL.md"},
                {"id": "DAKE_Mail_Draft", "title": "Mail Draft", "price": 500, "stripe_payment_link": "https://buy.stripe.com/old300", "source_original": "01_apps/AppPriceChange/ORIGINAL.md"},
                {"id": "dake_pdf_extract", "title": "Extract", "price": 500, "stripe_payment_link": None, "source_original": "01_apps/AppNew/ORIGINAL.md"},
            ]
            discontinued = (
                {"id": "dake_mail_kikuta", "title": "Kikuta", "stripe_payment_link": "https://buy.stripe.com/kikuta", "reason": "C"},
            )
            fake = FakeStripeClient(
                links=[
                    {"id": "plink_same", "url": "https://buy.stripe.com/same", "active": True, "metadata": {}, "after_completion": {"type": "hosted_confirmation"}},
                    {
                        "id": "plink_noop",
                        "url": "https://buy.stripe.com/noop",
                        "active": True,
                        "metadata": {"dake_item_id": "app_noop", "purchase_delivery_method": mod.DELIVERY_METHOD_R2},
                        "after_completion": {"type": "redirect", "redirect": {"url": mod.R2_REDIRECT_URL}},
                    },
                    {"id": "plink_old300", "url": "https://buy.stripe.com/old300", "active": True, "metadata": {"dake_item_id": "DAKE_Mail_Draft"}},
                    {"id": "plink_kikuta", "url": "https://buy.stripe.com/kikuta", "active": True, "metadata": {}},
                ],
                line_items_by_link={
                    "plink_same": [{"price": {"unit_amount": 500, "currency": "jpy", "product": "prod_same"}}],
                    "plink_noop": [{"price": {"unit_amount": 500, "currency": "jpy", "product": "prod_noop"}}],
                    "plink_old300": [{"price": {"unit_amount": 300, "currency": "jpy", "product": "prod_draft"}}],
                },
                products=[{"id": "prod_draft", "active": True, "metadata": {"dake_item_id": "DAKE_Mail_Draft"}}],
            )

            report = mod.reconcile_live(
                fake,  # type: ignore[arg-type]
                catalog,
                discontinued,
                apply=True,
                archive_discontinued=True,
                repo_root=repo_root,
            )

            s = report["summary"]
            self.assertEqual(s["errors"], 0)
            self.assertEqual(s["existing_links_updated"], 1)  # app_same
            self.assertEqual(s["noop"], 1)                    # app_noop
            self.assertEqual(s["new_links_created"], 2)       # DAKE_Mail_Draft + dake_pdf_extract
            self.assertEqual(s["replaced_old_links_archived"], 1)  # plink_old300
            self.assertEqual(s["discontinued_links_archived"], 1)  # plink_kikuta
            self.assertEqual(s["originals_updated"], 2)

            # Verify ORIGINAL.md write-backs
            draft_orig = (repo_root / "01_apps/AppPriceChange/ORIGINAL.md").read_text(encoding="utf-8")
            self.assertIn("https://buy.stripe.com/liveNew", draft_orig)
            extract_orig = (repo_root / "01_apps/AppNew/ORIGINAL.md").read_text(encoding="utf-8")
            self.assertIn("https://buy.stripe.com/liveNew", extract_orig)
            self.assertIn("- payment_status: stripe_ready", extract_orig)

            # Verify old link deactivated AFTER new link created
            deactivations = [c for c in fake.calls if c[0] == "update_payment_link" and c[2] == {"active": False}]
            self.assertEqual({c[1] for c in deactivations}, {"plink_old300", "plink_kikuta"})


if __name__ == "__main__":
    unittest.main()
