from __future__ import annotations

import math
import random
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pygame


APP_NAME = "Dakeしまりす不動産"
WINDOW_TITLE = "Dakeしまりす不動産"

WINDOW_WIDTH = 1180
WINDOW_HEIGHT = 820
FPS = 60

INITIAL_CASH = 20_000_000
INITIAL_REPUTATION = 50
MAX_DAYS = 180
BASE_MAINTENANCE_FEE = 50_000
LOW_CASH_WARNING = 3_000_000
DAILY_OPERATING_COST = 10_000

UI_TEXT = {
    "title": "Dakeしまりす不動産",
    "subtitle": "中古戸建を見極めて、買って、直して、売る。",
    "start": "はじめる",
    "new_case": "次の案件",
    "investigate": "調査する",
    "hearing": "ヒアリング",
    "buy": "買取する",
    "skip": "見送る",
    "renovate": "リフォーム",
    "sell_as_is": "現況販売",
    "list_sale": "販売開始",
    "next_day": "1日進める",
    "cash": "資金",
    "day": "日数",
    "profit": "利益",
    "reputation": "信用",
    "current_case": "現在の案件",
    "owned_property": "保有物件",
    "case_detail": "案件詳細",
    "cost_profit": "原価と利益見込み",
    "known_info": "わかっている情報",
    "hidden_info": "調査で見える情報",
    "status": "状態",
    "area": "エリア",
    "age": "築年数",
    "buy_price": "売主希望価格",
    "seller_asking_price": "売主希望価格",
    "purchase_price": "仕入価格",
    "appraisal_offer": "査定金額を提示",
    "appraisal_offer_title": "査定金額提示",
    "appraisal_price": "提示査定金額",
    "recommended_appraisal": "推奨査定価格",
    "recommended_after_check": "調査後に表示",
    "appraisal_reason": "査定理由",
    "appraisal_passed": "仕入契約がまとまりました。",
    "appraisal_rejected": "査定金額は通りませんでした。",
    "appraisal_rejected_note": "売主との条件が合わず、案件は終了しました。",
    "offer_at_asking": "希望価格のまま提示",
    "offer_minus_30": "30万円下げて提示",
    "offer_minus_50": "50万円下げて提示",
    "offer_minus_80": "80万円下げて提示",
    "offer_minus_100": "100万円下げて提示",
    "offer_risk_adjusted": "リスク反映価格で提示",
    "appraisal_warning": "調査前の提示は危険です。見えていないリスクを含んだ金額になります。",
    "appraisal_warning_hearing": "ヒアリング前の提示は危険です。売主しか知らない事情を見落としやすいです。",
    "appraisal_ok_reason": "調査と聞き取りを踏まえた金額です。",
    "appraisal_lowball_reason": "大きな問題が見えない中で低い提示です。",
    "appraisal_risk_reason": "判明したリスクを価格に反映しました。",
    "appraisal_high_risk_reason": "リスクに対して高めの仕入れです。",
    "operating_cost": "営業コスト",
    "daily_operating_cost": "営業コスト",
    "operating_cost_total": "営業コスト合計",
    "station_distance": "駅距離",
    "floor_plan": "間取り",
    "expected_sale": "想定販売価格",
    "source": "仕入れ経路",
    "parking": "駐車場",
    "corner_lot": "角地",
    "rebuildable": "再建築",
    "road_access": "接道",
    "incident": "事件事故",
    "flood": "浸水被害",
    "termite": "シロアリ",
    "illegal": "違反建築",
    "leftover": "残置物",
    "rain_leak": "雨漏り",
    "building_tilt": "建物傾き",
    "buyer_type": "買主層",
    "neighbor_garbage": "近隣ゴミ屋敷",
    "nearby_antisocial": "近隣反社拠点",
    "seller_antisocial": "売主反社関係",
    "risk_none": "なし",
    "risk_suspected": "疑い",
    "risk_confirmed": "あり",
    "risk_small": "少なめ",
    "risk_many": "多い",
    "legal_road_risk": "法務・道路",
    "building_risk": "建物",
    "neighborhood_risk": "周辺",
    "seller_risk": "売主",
    "cash_only": "現金客のみ",
    "loan_available": "ローン可",
    "loan_hard": "ローン難",
    "cash_center": "現金客中心",
    "yes": "あり",
    "no": "なし",
    "none": "なし",
    "reason": "理由",
    "possible": "可能",
    "not_possible": "不可",
    "good": "良好",
    "bad": "不良",
    "direct": "直接買取",
    "broker": "業者紹介",
    "candidate": "候補案件",
    "owned": "保有中",
    "renovating": "リフォーム中",
    "renovating_listed": "工事中販売",
    "listed": "販売中",
    "sold": "成約済み",
    "skipped": "見送り",
    "not_checked": "未確認",
    "checked": "確認済み",
    "survey_done": "調査済み",
    "hearing_done": "聞き取り済み",
    "days_held": "保有日数",
    "maintenance": "維持管理費",
    "brokerage_buy": "仕入手数料",
    "renovation_cost": "リフォーム費",
    "repair_cost": "追加修繕費",
    "total_cost": "総原価",
    "listed_price": "販売価格",
    "sale_strategy": "販売方針",
    "sale_chance": "本日の売却見込み",
    "target_profit": "目標利益",
    "market_feel": "市場感",
    "select_renovation": "リフォーム方針を選んでください",
    "select_price": "販売価格を選んでください",
    "advance_to_finish": "リフォーム完了まで日を進めましょう",
    "listed_waiting": "販売中です。日を進めると売却判定します",
    "sold_result": "成約結果",
    "result": "180日結果",
    "final_cash": "最終資金",
    "cash_gain": "純増資金",
    "cash_memo": "資金メモ",
    "cash_purchase_note": "現金で仕入れました。",
    "cash_sale_note": "売却時に販売代金が戻ります。",
    "cash_after_sale_note": "売却代金が資金に戻りました。",
    "cash_delta": "資金増減",
    "sold_count": "成約件数",
    "skipped_count": "見送り件数",
    "best_profit": "最大利益案件",
    "restart": "もう一度",
    "quit": "終了",
    "no_property": "案件がありません",
    "insufficient_cash": "資金が足りません",
    "click_start": "はじめるを押して案件を見に行きましょう。",
    "auto_customer": "自社客",
    "outside_customer": "外部仲介客",
    "sale_channel": "売却経路",
    "brokerage_income": "自社仲介収入",
    "outside_fee": "外部仲介手数料",
    "net_proceeds": "売却手取",
    "sale_price": "販売価格",
    "confirmed_profit": "確定利益",
    "projected_profit": "見込利益",
    "target_gap": "目標までの差額",
    "after_strategy": "方針選択後",
    "after_price": "販売価格選択後",
    "sold_banner": "成約しました！",
    "price_early": "早期売却価格",
    "price_fair": "適正価格",
    "price_strong": "強気価格",
    "as_is": "現況販売",
    "light_renovation": "軽めリフォーム",
    "standard_renovation": "標準リフォーム",
    "full_renovation": "しっかり再生",
    "shimarisu_name": "しまりす君",
    "office": "しまりす事務所",
    "memo": "メモ",
    "remaining_work": "残り工期",
    "best_profit_none": "まだ成約なし",
    "renovation_done_comment": "リフォーム完了です。販売価格を決めましょう。",
    "log_new_case": "新規案件: {name}",
    "log_investigate": "調査で法的・物理的リスクを確認しました。",
    "log_hearing": "ヒアリングで売主側の事情を確認しました。",
    "log_purchase": "買取成立: {amount} を支払いました。",
    "log_purchase_event": "買取後イベント: {message}",
    "log_extra_cost": "追加費用: {amount}",
    "log_skip": "見送り: {name}",
    "log_as_is": "現況販売の方針にしました。",
    "log_renovation": "{label}: {amount} / {days}日",
    "log_renovation_done": "リフォームが完了しました。",
    "log_listed": "販売開始: {label} {amount}",
    "log_maintenance": "維持管理費: {amount}",
    "log_sold": "成約: {amount} / 利益 {profit}",
    "log_cash_short": "資金が足りません: 必要 {amount}",
    "log_cash_short_renovation": "資金が足りません: リフォーム {amount}",
    "event_flood_found": "浸水履歴が見つかりました",
    "event_termite_found": "シロアリ被害が見つかりました",
    "event_leftover_found": "残置物の処分が必要でした",
    "event_illegal_found": "違反建築が判明しました",
    "event_rebuild_found": "再建築不可が判明しました",
    "event_road_found": "接道条件が弱いことが判明しました",
    "event_incident_found": "事件事故の告知事項が見つかりました",
    "event_rain_leak_found": "雨漏りの痕跡が見つかりました",
    "event_tilt_found": "建物の傾きが見つかりました",
    "event_garbage_found": "近隣ゴミ屋敷が見つかりました",
    "event_nearby_antisocial_found": "近隣反社拠点が見つかりました",
    "event_seller_antisocial_found": "売主反社関係が判明しました",
    "confirm_next": "結果を確認して次へ",
    "confirm_continue": "確認して続ける",
    "skipped_done": "見送り済み",
    "skip_cost_note": "案件探しに2日使いました。",
    "log_skip_cost": "案件探しで2日経過しました。",
    "log_broker_skip_rep": "業者紹介を見送りました。評判 -1",
    "good_skip_comment": "見送る判断も仕事です。",
    "investigate_hint": "調査：1日使って、法的・物理的リスクを確認します。",
    "hearing_hint": "ヒアリング：1日使って、売主しか知らない事情を確認します。",
    "risk_notice": "隠れリスクが発覚しました",
    "risk_investigate_short": "調査不足でした。{message}",
    "risk_hearing_short": "ヒアリング不足でした。{message}",
    "event_neighbor_found": "近隣トラブルが見つかりました",
    "neighbor_trouble": "近隣トラブル",
    "sale_feedback": "販売反応",
    "feedback_price_high": "価格がやや強気です。",
    "feedback_strategy_strong": "強気価格のため、様子見されています。",
    "feedback_rebuild": "再建築不可のため、問い合わせが少なめです。",
    "feedback_road": "接道条件が弱く、出口が絞られています。",
    "feedback_illegal": "違反建築のため、現金客に限られます。",
    "feedback_incident": "事件事故の印象で反応が鈍いです。",
    "feedback_flood": "浸水履歴を気にする買主がいます。",
    "feedback_termite": "シロアリ被害が不安材料です。",
    "feedback_parking": "駐車場なしが弱点です。",
    "feedback_old": "築古のため、現況確認が慎重です。",
    "feedback_long": "長期在庫で反響が弱くなっています。",
    "feedback_reputation": "信用が低く、問い合わせが伸びません。",
    "feedback_cash_center": "現金客中心のため、成約まで時間がかかりそうです。",
    "feedback_normal": "大きな弱点はありません。反響待ちです。",
    "price_down_500": "価格を50万円下げる",
    "price_down_1000": "価格を100万円下げる",
    "log_price_down": "販売価格を{amount}下げました。問い合わせは少し増えそうです。",
    "price_down_blocked": "これ以上下げると、利益がほとんど残りません。",
    "price_down_blocked_extra": "損切り処分も検討する場面です。",
    "cash_assets": "資金と資産",
    "cash_on_hand": "現金",
    "holding_value": "保有物件評価",
    "total_assets": "総資産目安",
    "asset_gain": "純増資産",
    "unrealized_profit": "含み損益",
    "unsold_count": "売れ残り件数",
    "cash_warning": "資金が少なくなっています。在庫を抱えすぎると動けなくなります。",
    "cash_short_result": "資金ショートです。売れるまで耐える必要がありました。",
    "case_memo": "案件メモ",
    "rough_profit": "想定粗利",
    "investigation_state": "調査状態",
    "hearing_state": "ヒアリング状態",
    "attention_points": "注意ポイント",
    "skip_result": "見送り結果",
    "skip_message": "この案件は見送りました。",
    "skip_reason": "見送り理由",
    "skip_reason_risk": "リスク高",
    "skip_reason_margin": "利幅不足",
    "skip_reason_cash": "資金温存",
    "skip_reason_optional": "任意",
    "elapsed_days": "経過日数",
    "credit_delta": "信用変化",
    "credit_reason": "理由",
    "trust_evaluation": "信用評価",
    "loss_count": "赤字件数",
    "target_missed_sales": "目標未達成約",
    "uninvestigated_sales": "未調査成約",
    "unheard_sales": "未ヒアリング成約",
    "market_freshness": "市場感",
    "sales_days": "販売日数",
    "confirmed_loss": "確定損失",
    "sale_negotiation": "成約時調整",
    "feedback_under_investigated": "調査不足で出口が読みづらいです。",
    "feedback_under_heard": "聞き取り不足で買主説明に不安があります。",
    "feedback_market_stale": "販売開始から時間が経ち、市場感が落ちています。",
    "feedback_suburban_one_parking": "駐車場1台なので、郊外ではやや弱いです。",
    "feedback_price_floor": "これ以上の値下げは利益がほとんど残りません。損切り処分も検討する場面です。",
    "feedback_station_far": "駅から遠く、買主層が絞られます。",
    "feedback_floor_2ldk": "2LDKなので、家族向けとしては少し弱いです。",
    "feedback_floor_old": "古い間取り感があり、リフォーム効果が問われます。",
    "feedback_purchase_heavy": "仕入価格が重く、利益を残しにくいです。",
    "feedback_risk_purchase_heavy": "リスクに対して高く買っているため、出口が厳しくなっています。",
    "feedback_parking_floor_mismatch": "間取りに対して駐車場が足りません。",
    "feedback_family_parking": "家族向けですが、車2台需要に弱いです。",
    "feedback_early_big_drop": "販売直後の大幅値下げで、市場に弱気が伝わっています。",
    "feedback_many_price_drops": "短期間の複数値下げで、買主が指値しやすくなっています。",
    "feedback_need_price_review": "60日超で価格見直しがなく、反響が細っています。",
    "feedback_rain_leak": "雨漏り懸念で、建物説明が重くなっています。",
    "feedback_tilt": "建物の傾きで、買主の不安が大きいです。",
    "feedback_neighbor_garbage": "近隣ゴミ屋敷で、内見後の反応が落ちます。",
    "feedback_nearby_antisocial": "近隣反社拠点が重く、出口が大きく狭まります。",
    "sold_loss_reason": "値下げと維持管理費が重なりました。",
    "credit_full_good": "調査とヒアリングを行い、納得感のある成約でした。",
    "credit_partial_good": "確認は一部に留まりましたが、黒字でまとめました。",
    "credit_no_check_good": "利益は出ましたが、調査不足のまま進めたため信用は伸びませんでした。",
    "credit_hidden_risk": "隠れリスク発覚後の成約です。説明責任の重さが残りました。",
    "credit_loss": "赤字成約です。仕入れ判断と価格調整を見直しましょう。",
    "credit_target_missed": "黒字ではありますが、目標利益には届きませんでした。",
    "credit_under_construction_discount": "工事中で早く売れましたが、指値を受けています。出口判断は慎重に見ましょう。",
    "sale_result_target_clear": "目標利益を超えました。出口設計がうまくいきました。",
    "sale_result_reform_thin": "売れましたが、リフォーム再販としては利益が薄いです。250万円を切るなら、仕入れか出口の見直しが必要です。",
    "sale_result_reform_loss": "赤字です。仕入れ価格か出口判断を見直しましょう。",
    "sale_result_as_is_clear": "現況販売としては最低ラインを超えました。",
    "sale_result_as_is_thin": "売れましたが、現況販売としては利益が薄いです。",
    "sale_result_black_target_missed": "黒字ではありますが、成功とは言い切れません。",
    "trust_comment_good": "堅実です。調査と説明を重ね、次につながる仕事でした。",
    "trust_comment_mixed": "利益は出ていますが、調査不足の案件が目立ちました。",
    "trust_comment_bad": "資金と信用の両方が傷んでいます。仕入れ判断を見直しましょう。",
    "trust_comment_target_missed": "黒字は出ていますが、目標利益未達の成約が目立ちました。",
    "trust_comment_thin_exit": "売り切っていますが、利益の薄い出口が多いです。",
    "trust_comment_cash_recovered": "資金回収はできています。次は仕入れ価格をもう少し絞りましょう。",
    "exit_method": "出口",
    "old_house_land_exit": "現況販売 / 古家付き売地",
    "purchase_banned": "買付禁止",
    "purchase_banned_skip": "買付禁止：見送る",
    "purchase_banned_notice": "これは利益ではなく撤退です。買わないことが仕事です。",
    "renovation_listing": "工事中販売を開始しました。",
    "start_renovation_listing": "工事中販売を開始",
    "under_construction_offer": "工事中申込",
    "offer_received": "購入申込が入りました。",
    "offer_price": "申込価格",
    "offer_discount": "指値",
    "offer_reason": "指値理由",
    "accept_offer": "指値を受ける",
    "reject_offer": "断って販売継続",
    "full_price_offer": "満額申込",
    "offer_reason_full": "条件が良く、満額でまとまりそうです。",
    "offer_reason_stale": "販売日数が長く、買主が弱気です。",
    "offer_reason_risk": "リスク説明を理由に指値されています。",
    "offer_reason_cash": "現金決済の早さを理由にした指値です。",
    "offer_reason_rounding": "端数調整の範囲です。利益目標を確認しましょう。",
    "offer_reason_million": "100万円の指値です。飲むなら理由が必要です。",
    "log_offer_received": "購入申込: {amount} / 指値 {discount}",
    "log_offer_accepted": "指値を受けました: {amount}",
    "log_offer_rejected": "申込を断って販売継続しました。",
    "offer_accepted_credit": "この指値なら、資金回収とのバランスで検討できます。",
    "offer_loss_credit": "指値を受けて赤字になりました。出口判断を見直しましょう。",
    "offer_rejected": "申込を断りました。次の買主がいるかを見ます。",
    "quick_exit": "早い資金回収でした。",
    "slow_exit": "資金回収に時間がかかりました。",
    "loss_cut": "損切り処分",
    "loss_cut_result": "損切り処分しました",
    "loss_cut_price": "処分価格",
    "loss_cut_channel": "損切り処分",
    "log_loss_cut": "損切り処分: {amount} / 利益 {profit}",
    "credit_loss_cut": "痛いですが、次へ進むための出口を作りました。",
    "unsold_evaluation": "売れ残り評価",
    "unsold_has_inventory": "売れ残りがあります。仕入判断と出口設計を見直しましょう。",
    "unsold_clear": "売れ残りなし。出口まで作れています。",
    "unsold_credit_change": "信用変化",
    "unsold_penalty_inventory": "売れ残り",
    "unsold_penalty_stale": "市場感低下",
    "unsold_penalty_unrealized_loss": "含み損",
    "unsold_penalty_low_cash": "資金不足",
    "result_unsold_asset_plus": "資産は残っていますが、売り切れていません。次は出口まで見て仕入れましょう。",
    "result_unsold_cash_short": "在庫に資金が寝ています。売れなければ、次の仕入れに進めません。",
    "result_unsold_loss": "仕入れ失敗です。損切りも含めて、出口を作る必要がありました。",
}

SHIMARISU_COMMENTS = {
    "title": "安いには、理由がありそうです。数字と出口を見ていきましょう。",
    "new_case": "新しい案件です。見える情報だけで飛びつかないのが大事です。",
    "investigated": "法的なところが見えてきました。出口の広さを確認しましょう。",
    "hearing": "聞いていないことが、あとから出てくることがあります。",
    "bought": "握りました。ここからは原価をふくらませすぎない勝負です。",
    "appraisal_mode": "査定金額を一度だけ提示します。調査と聞き取りの材料で勝負しましょう。",
    "appraisal_passed": "仕入契約がまとまりました。ここから出口設計の責任が始まります。",
    "appraisal_rejected": "売主と条件が合いませんでした。次の案件に切り替えましょう。",
    "skipped": "この案件は、握らない強さかもしれません。",
    "cash_short": "資金が足りません。在庫と現金を少し見ましょう。",
    "renovate_light": "軽めでも、第一印象はけっこう変わります。",
    "renovate_standard": "350万円で、だいぶ見え方が変わりそうです。",
    "renovate_full": "しっかり直す分、売値と日数のバランスを見ましょう。",
    "as_is": "現況で売る判断も、悪くなさそうです。",
    "listed": "販売開始です。価格が強いと、時間がかかるかもしれません。",
    "sold_good": "いい出口でした。次もこの調子でいきましょう。",
    "sold_bad": "売れましたが、利益は薄めです。次は仕入れを締めましょう。",
    "not_sold": "今日は反響止まりでした。焦らず見ます。",
    "long_hold": "在庫が長くなってきました。少し資金を見ましょう。",
    "illegal": "ローンが使えないので、現金客向けになりそうです。",
    "parking": "駐車場2台は強いです。",
    "wide_parking_bad": "家は広いですが、駐車場が足りません。ここは出口で響きます。",
    "rebuild_bad": "再建築不可です。出口はかなり絞られます。",
    "young": "築浅ですが、利益を乗せる余白が少ないです。",
    "hidden": "聞いていないことが、あとから出てきました。",
    "risk_found": "確認不足が響きました。ここは一度、結果を見て判断しましょう。",
    "skip_good": "見送る判断も仕事です。",
    "price_down": "値下げしました。利益は減りますが、反響は少し増えそうです。",
    "price_down_blocked": "これ以上下げると、利益がほとんど残りません。",
    "renovating_offer": "工事中でも申込が入るのは、出口が見えている証拠です。",
    "offer_received": "購入申込です。指値を受ける理由があるか、次の買主がいるかを見ましょう。",
    "offer_accept": "痛い指値でも、資金を戻して次へ進めるなら出口です。",
    "offer_reject": "断るなら、次の買主がいる根拠を見ましょう。",
    "lowball_market": "安い理由を買主も見ています。価格の下げ方には注意です。",
    "risk_retreat": "これは利益ではなく撤退です。買わないことが仕事です。",
    "loss_cut": "痛いですが、次へ進むための出口です。",
    "low_cash": "資金が少なくなっています。在庫を抱えすぎると動けなくなります。",
    "result_good": "堅く増やせました。小さく作って、ちゃんと回っています。",
    "result_even": "大きくは伸びませんでしたが、ループは回せています。",
    "result_bad": "次は仕入れ前の確認を厚めにしましょう。",
}

COLORS = {
    "bg": (245, 247, 244),
    "panel": (255, 255, 252),
    "panel_alt": (249, 251, 247),
    "line": (214, 221, 211),
    "line_dark": (166, 176, 160),
    "text": (38, 43, 40),
    "muted": (101, 112, 102),
    "soft": (232, 238, 228),
    "soft_green": (218, 235, 220),
    "green": (82, 139, 94),
    "green_dark": (52, 105, 66),
    "orange": (222, 139, 75),
    "red": (190, 76, 70),
    "yellow": (248, 225, 154),
    "white": (255, 255, 255),
    "shadow": (217, 222, 216),
    "button": (255, 255, 255),
    "button_hover": (232, 243, 233),
    "button_disabled": (231, 233, 230),
}

JAPANESE_FONT_NAMES = (
    "BIZ UDPGothic",
    "Yu Gothic UI",
    "Meiryo",
    "Noto Sans CJK JP",
    "MS Gothic",
)

AREAS = (
    "千葉市若葉区",
    "佐倉市",
    "四街道市",
    "八千代市",
    "印西市",
    "船橋市郊外",
    "成田市",
    "市原市",
)

PROPERTY_NAMES = (
    "角地の戸建",
    "森のそばの戸建",
    "小さな再生案件",
    "築古戸建",
    "駅遠めの戸建",
    "庭付き中古戸建",
    "駐車場広めの家",
    "日当たりの良い家",
)

FLOOR_PLANS = ("2LDK", "3LDK", "4LDK", "5DK", "5LDK", "4DK")

RISK_LABELS = {
    "rebuildable": UI_TEXT["rebuildable"],
    "road_access_good": UI_TEXT["road_access"],
    "incident_property": UI_TEXT["incident"],
    "flood_damage": UI_TEXT["flood"],
    "termite_damage": UI_TEXT["termite"],
    "illegal_building": UI_TEXT["illegal"],
    "leftover_items": UI_TEXT["leftover"],
    "rain_leak_level": UI_TEXT["rain_leak"],
    "building_tilt_level": UI_TEXT["building_tilt"],
    "neighbor_trouble": UI_TEXT["neighbor_trouble"],
    "neighbor_garbage_house": UI_TEXT["neighbor_garbage"],
    "nearby_antisocial_base": UI_TEXT["nearby_antisocial"],
    "seller_antisocial": UI_TEXT["seller_antisocial"],
}

RENOVATION_OPTIONS = (
    {"label_key": "as_is", "cost": 0, "days": 0},
    {"label_key": "light_renovation", "cost": 1_500_000, "days": 7},
    {"label_key": "standard_renovation", "cost": 3_500_000, "days": 14},
    {"label_key": "full_renovation", "cost": 5_000_000, "days": 28},
)

SALE_OPTIONS = (
    {"label_key": "price_early", "key": "early", "profit_delta": -500_000},
    {"label_key": "price_fair", "key": "fair", "profit_delta": 0},
    {"label_key": "price_strong", "key": "strong", "profit_delta": 1_000_000},
)


@dataclass
class PropertyCase:
    id: int
    name: str
    area: str
    buy_price: int
    expected_sale_price: int
    source: str
    age: int
    station_distance_m: int
    station_walk_minutes: int
    floor_plan: str
    parking_count: int
    corner_lot: bool
    rebuildable: bool
    road_access_good: bool
    incident_property: bool
    flood_damage: bool
    termite_damage: bool
    illegal_building: bool
    leftover_items: bool
    neighbor_trouble: bool
    rain_leak_level: str = "none"
    building_tilt_level: str = "none"
    termite_level: str = "none"
    leftover_level: str = "none"
    neighbor_garbage_house: bool = False
    nearby_antisocial_base: bool = False
    seller_antisocial_level: str = "none"
    seller_antisocial: bool = False
    rebuild_blocker_reason: str = ""
    status: str = "candidate"
    known_flags: dict[str, bool] = field(default_factory=dict)
    investigated: bool = False
    heard_from_seller: bool = False
    renovation_budget: int = 0
    renovation_days_left: int = 0
    listed_price: int = 0
    days_held: int = 0
    maintenance_cost_total: int = 0
    brokerage_fee_buy: int = 0
    repair_cost_total: int = 0
    seller_asking_price: int = 0
    appraisal_price: int = 0
    appraisal_result: str = ""
    appraisal_reason: str = ""
    appraisal_credit_delta: int = 0
    appraisal_mode: bool = False
    overpaid_purchase: bool = False
    renovation_label_key: str = ""
    sale_strategy_key: str = ""
    target_profit: int = 0
    sold_price: int = 0
    sold_profit: int = 0
    sale_channel: str = ""
    brokerage_income: int = 0
    outside_fee: int = 0
    net_proceeds: int = 0
    last_sale_feedback: list[str] = field(default_factory=list)
    price_cut_count: int = 0
    price_drop_total: int = 0
    early_big_price_drop: bool = False
    many_early_price_drops: bool = False
    listed_days: int = 0
    market_freshness: int = 100
    hidden_risk_found: bool = False
    credit_delta: int = 0
    credit_reason: str = ""
    sale_result_comment: str = ""
    sale_negotiation_discount: int = 0
    disposal_sale: bool = False
    purchase_offer_price: int = 0
    purchase_offer_discount: int = 0
    purchase_offer_reason: str = ""
    purchase_offer_self_buyer: bool = True
    sold_under_construction: bool = False
    banned_skip: bool = False


@dataclass
class Button:
    rect: pygame.Rect
    text: str
    action: str
    payload: Any = None
    enabled: bool = True


def resource_path(relative: str) -> Path:
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parent
    return base / relative


def calc_brokerage_fee(price: int) -> int:
    if price <= 8_000_000:
        return 330_000
    return int((price * 0.03 + 60_000) * 1.10)


def yen(value: int) -> str:
    return f"{value:,}円"


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def weighted_choice(items: tuple[Any, ...], weights: tuple[int, ...]) -> Any:
    return random.choices(items, weights=weights, k=1)[0]


def fit_surface(surface: pygame.Surface, size: tuple[int, int]) -> pygame.Surface:
    target_w, target_h = size
    src_w, src_h = surface.get_size()
    scale = min(target_w / src_w, target_h / src_h)
    new_size = (max(1, int(src_w * scale)), max(1, int(src_h * scale)))
    return pygame.transform.smoothscale(surface, new_size)


class ShimarisuRealEstateGame:
    def __init__(self) -> None:
        pygame.init()
        pygame.display.set_caption(WINDOW_TITLE)
        self.screen = pygame.display.set_mode((WINDOW_WIDTH, WINDOW_HEIGHT))
        self.clock = pygame.time.Clock()
        self.fonts: dict[tuple[int, bool], pygame.font.Font] = {}

        self.shimarisu = self.load_first_image(("assets/shimarisu_transparent.png", "assets/shimarisu.png"))
        self.shimarisu_title = self.load_first_image(
            ("assets/shimarisu_transparent.png", "assets/shimarisu_title.png", "assets/shimarisu.png")
        )
        self.village = self.load_image("assets/village.png")

        self.screen_name = "title"
        self.buttons: list[Button] = []
        self.mouse_pos = (0, 0)
        self.running = True

        self.reset_game()

    def reset_game(self) -> None:
        self.day = 1
        self.cash = INITIAL_CASH
        self.reputation = INITIAL_REPUTATION
        self.case_id = 0
        self.current_property: PropertyCase | None = None
        self.sold_count = 0
        self.skipped_count = 0
        self.loss_count = 0
        self.target_missed_sales = 0
        self.uninvestigated_sales = 0
        self.unheard_sales = 0
        self.operating_cost_total = 0
        self.best_profit = -10**12
        self.best_profit_name = ""
        self.message_log: list[str] = []
        self.pending_confirmation: str | None = None
        self.result_assessed = False
        self.result_inventory_penalties: list[tuple[str, int]] = []
        self.comment = SHIMARISU_COMMENTS["title"]

    def load_first_image(self, relatives: tuple[str, ...]) -> pygame.Surface | None:
        for relative in relatives:
            image = self.load_image(relative)
            if image:
                return image
        return None

    def load_image(self, relative: str) -> pygame.Surface | None:
        path = resource_path(relative)
        try:
            return pygame.image.load(path).convert_alpha()
        except (pygame.error, OSError):
            return None

    def get_font(self, size: int, bold: bool = False) -> pygame.font.Font:
        key = (size, bold)
        if key in self.fonts:
            return self.fonts[key]

        for name in JAPANESE_FONT_NAMES:
            try:
                font_path = pygame.font.match_font(name, bold=bold)
                if font_path:
                    font = pygame.font.Font(font_path, size)
                    self.fonts[key] = font
                    return font
                font = pygame.font.SysFont(name, size, bold=bold)
                if font:
                    self.fonts[key] = font
                    return font
            except pygame.error:
                continue

        font = pygame.font.Font(None, size)
        self.fonts[key] = font
        return font

    def run(self) -> None:
        while self.running:
            self.clock.tick(FPS)
            self.mouse_pos = pygame.mouse.get_pos()
            for event in pygame.event.get():
                self.handle_event(event)
            self.draw()
            pygame.display.flip()
        pygame.quit()

    def handle_event(self, event: pygame.event.Event) -> None:
        if event.type == pygame.QUIT:
            self.running = False
            return
        if event.type == pygame.MOUSEBUTTONDOWN and event.button == 1:
            for button in self.buttons:
                if button.enabled and button.rect.collidepoint(event.pos):
                    self.perform_action(button.action, button.payload)
                    return

    def perform_action(self, action: str, payload: Any = None) -> None:
        if action == "start":
            self.screen_name = "main"
            self.generate_new_case()
        elif action == "restart":
            self.reset_game()
            self.screen_name = "title"
        elif action == "quit":
            self.running = False
        elif action == "new_case":
            self.generate_new_case()
        elif action == "investigate":
            self.investigate()
        elif action == "hearing":
            self.hearing()
        elif action == "buy":
            self.buy_property()
        elif action == "appraisal_start":
            self.start_appraisal_offer()
        elif action == "appraisal_offer":
            self.submit_appraisal_offer(int(payload))
        elif action == "skip":
            self.skip_property()
        elif action == "purchase_banned_skip":
            self.skip_banned_property()
        elif action == "confirm_next":
            self.confirm_and_go_next()
        elif action == "confirm_continue":
            self.pending_confirmation = None
        elif action == "renovation":
            self.choose_renovation(payload)
        elif action == "list_price":
            self.choose_list_price(payload)
        elif action == "discount":
            self.discount_listed_price(int(payload))
        elif action == "loss_cut":
            self.loss_cut_property()
        elif action == "accept_offer":
            self.accept_purchase_offer()
        elif action == "reject_offer":
            self.reject_purchase_offer()
        elif action == "next_day":
            self.advance_day()

    def generate_new_case(self) -> None:
        self.pending_confirmation = None
        self.case_id += 1
        self.current_property = self.create_property(self.case_id)
        self.comment = self.initial_comment(self.current_property)
        self.add_log(UI_TEXT["log_new_case"].format(name=self.current_property.name))

    def confirm_and_go_next(self) -> None:
        self.pending_confirmation = None
        prop = self.current_property
        if self.cash < LOW_CASH_WARNING and prop and prop.status in {"owned", "renovating", "listed"}:
            self.comment = SHIMARISU_COMMENTS["low_cash"]

        if self.day > MAX_DAYS and not self.pending_confirmation:
            self.finish_game()
            return
        self.generate_new_case()

    def create_property(self, case_id: int) -> PropertyCase:
        age = weighted_choice(
            tuple(range(4, 56)),
            tuple(4 if y < 15 else 8 if y < 35 else 10 for y in range(4, 56)),
        )
        area = random.choice(AREAS)
        name = f"{area} {random.choice(PROPERTY_NAMES)}"
        distance_pattern = random.choices(("near", "normal", "far", "very_far"), weights=(22, 44, 24, 10), k=1)[0]
        if distance_pattern == "near":
            station_distance_m = random.randint(300, 800)
        elif distance_pattern == "normal":
            station_distance_m = random.randint(900, 1600)
        elif distance_pattern == "far":
            station_distance_m = random.randint(1700, 2200)
        else:
            station_distance_m = random.randint(2300, 3000)
        station_walk_minutes = math.ceil(station_distance_m / 80)
        floor_plan = random.choices(FLOOR_PLANS, weights=(16, 32, 26, 10, 8, 8), k=1)[0]

        if age < 15:
            buy_price = random.randint(12_000_000, 18_500_000)
            expected = buy_price + random.randint(1_800_000, 4_200_000)
        elif age < 35:
            buy_price = random.randint(8_000_000, 15_500_000)
            expected = buy_price + random.randint(2_500_000, 6_500_000)
        else:
            buy_price = random.randint(4_500_000, 11_500_000)
            expected = buy_price + random.randint(3_500_000, 8_500_000)

        parking_count = weighted_choice((0, 1, 2, 3), (15, 46, 30, 9))
        corner_lot = random.random() < 0.22
        road_access_good = random.random() > (0.08 if age < 35 else 0.17)
        rebuild_blocker_reason = ""
        if not road_access_good:
            rebuildable = False
            rebuild_blocker_reason = random.choice(("接道2m未満", "建築基準法道路ではない"))
        else:
            has_other_blocker = random.random() < (0.04 if age < 35 else 0.12)
            rebuildable = not has_other_blocker
            if has_other_blocker:
                rebuild_blocker_reason = random.choice(
                    ("市街化調整区域", "法令制限あり", "不明な制限あり")
                )
        illegal_building = random.random() < (0.05 if age < 25 else 0.14)
        flood_damage = random.random() < (0.07 if area in {"佐倉市", "市原市"} else 0.05)
        termite_level = weighted_choice(
            ("none", "suspected", "confirmed"),
            (96, 3, 1) if age < 25 else (82, 10, 8),
        )
        termite_damage = termite_level != "none"
        incident_property = random.random() < 0.06
        leftover_level = weighted_choice(
            ("none", "small", "many"),
            (90, 7, 3) if age < 25 else (76, 14, 10),
        )
        leftover_items = leftover_level != "none"
        neighbor_trouble = random.random() < (0.04 if age < 25 else 0.10)
        rain_leak_level = weighted_choice(
            ("none", "suspected", "confirmed"),
            (86, 10, 4) if age < 25 else (72, 18, 10),
        )
        building_tilt_level = weighted_choice(
            ("none", "suspected", "confirmed"),
            (94, 5, 1) if age < 25 else (84, 12, 4),
        )
        neighbor_garbage_house = random.random() < (0.04 if age < 25 else 0.09)
        nearby_antisocial_base = random.random() < 0.045
        seller_antisocial_level = weighted_choice(("none", "suspected", "confirmed"), (96, 3, 1))
        seller_antisocial = seller_antisocial_level == "confirmed"
        source = "direct" if random.random() < 0.46 else "broker"

        if parking_count >= 2:
            expected += 800_000
        if corner_lot:
            expected += 500_000
        if not rebuildable:
            expected -= 2_000_000
            buy_price -= 900_000
        if illegal_building:
            expected -= 1_200_000
            buy_price -= 600_000
        if incident_property:
            expected -= 1_000_000
            buy_price -= 700_000
        if rain_leak_level == "suspected":
            expected -= 600_000
            buy_price -= 250_000
        elif rain_leak_level == "confirmed":
            expected -= 1_300_000
            buy_price -= 600_000
        if building_tilt_level == "suspected":
            expected -= 900_000
            buy_price -= 350_000
        elif building_tilt_level == "confirmed":
            expected -= 2_000_000
            buy_price -= 900_000
        if neighbor_garbage_house:
            expected -= 600_000
            buy_price -= 250_000
        if nearby_antisocial_base:
            expected -= 2_200_000
            buy_price -= 1_000_000
        if seller_antisocial:
            buy_price -= 700_000
        if station_walk_minutes <= 10:
            expected += 700_000
            buy_price += 300_000
        elif station_walk_minutes > 30:
            expected -= 1_400_000
            buy_price -= 800_000
        elif station_walk_minutes > 20:
            expected -= 800_000
            buy_price -= 400_000
        if parking_count >= 2 and station_walk_minutes > 20:
            expected += 400_000
        if floor_plan in {"3LDK", "4LDK"}:
            expected += 500_000
            buy_price += 200_000
        elif floor_plan == "2LDK":
            expected -= 500_000
            buy_price -= 200_000
        elif floor_plan in {"5DK", "5LDK", "4DK"}:
            expected -= 300_000
        if floor_plan in {"4LDK", "5DK", "5LDK"} and parking_count == 0:
            expected -= 900_000
            buy_price -= 300_000
        elif floor_plan in {"4LDK", "5DK", "5LDK"} and parking_count == 1:
            expected -= 400_000

        buy_price = max(3_800_000, round(buy_price / 100_000) * 100_000)
        expected = max(buy_price + 1_000_000, round(expected / 100_000) * 100_000)

        return PropertyCase(
            id=case_id,
            name=name,
            area=area,
            buy_price=buy_price,
            expected_sale_price=expected,
            source=source,
            age=age,
            station_distance_m=station_distance_m,
            station_walk_minutes=station_walk_minutes,
            floor_plan=floor_plan,
            parking_count=parking_count,
            corner_lot=corner_lot,
            rebuildable=rebuildable,
            road_access_good=road_access_good,
            incident_property=incident_property,
            flood_damage=flood_damage,
            termite_damage=termite_damage,
            illegal_building=illegal_building,
            leftover_items=leftover_items,
            neighbor_trouble=neighbor_trouble,
            rain_leak_level=rain_leak_level,
            building_tilt_level=building_tilt_level,
            termite_level=termite_level,
            leftover_level=leftover_level,
            neighbor_garbage_house=neighbor_garbage_house,
            nearby_antisocial_base=nearby_antisocial_base,
            seller_antisocial_level=seller_antisocial_level,
            seller_antisocial=seller_antisocial,
            rebuild_blocker_reason=rebuild_blocker_reason,
            seller_asking_price=buy_price,
            known_flags={},
        )

    def initial_comment(self, prop: PropertyCase) -> str:
        if self.has_wide_parking_mismatch(prop):
            return SHIMARISU_COMMENTS["wide_parking_bad"]
        if prop.parking_count >= 2:
            return SHIMARISU_COMMENTS["parking"]
        if prop.age < 15:
            return SHIMARISU_COMMENTS["young"]
        return SHIMARISU_COMMENTS["new_case"]

    def add_log(self, message: str) -> None:
        self.message_log.insert(0, message)
        self.message_log = self.message_log[:6]

    def apply_operating_cost(self, days: int = 1) -> None:
        if days <= 0:
            return
        amount = DAILY_OPERATING_COST * days
        self.cash -= amount
        self.operating_cost_total += amount

    def investigate(self) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate":
            return
        if prop.known_flags.get("investigated"):
            return
        prop.known_flags["investigated"] = True
        prop.investigated = True
        for key in (
            "rebuildable",
            "road_access_good",
            "illegal_building",
            "flood_damage",
            "rain_leak_level",
            "building_tilt_level",
        ):
            prop.known_flags[key] = True
        self.comment = SHIMARISU_COMMENTS["investigated"]
        if not prop.rebuildable:
            self.comment = SHIMARISU_COMMENTS["rebuild_bad"]
        if prop.illegal_building:
            self.comment = SHIMARISU_COMMENTS["illegal"]
        if prop.building_tilt_level != "none":
            self.comment = SHIMARISU_COMMENTS["lowball_market"]
        self.add_log(UI_TEXT["log_investigate"])
        self.advance_day(allow_sale=False)

    def hearing(self) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate":
            return
        if prop.known_flags.get("hearing"):
            return
        prop.known_flags["hearing"] = True
        prop.heard_from_seller = True
        for key in (
            "incident_property",
            "termite_damage",
            "leftover_items",
            "neighbor_trouble",
            "neighbor_garbage_house",
            "nearby_antisocial_base",
            "seller_antisocial",
        ):
            prop.known_flags[key] = True
        self.comment = SHIMARISU_COMMENTS["risk_retreat"] if self.is_purchase_banned(prop) else SHIMARISU_COMMENTS["hearing"]
        self.add_log(UI_TEXT["log_hearing"])
        self.advance_day(allow_sale=False)

    def buy_property(self) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate":
            return
        if self.is_purchase_banned(prop):
            self.skip_banned_property()
            return
        self.complete_purchase(prop.seller_asking_price or prop.buy_price, UI_TEXT["appraisal_ok_reason"])

    def start_appraisal_offer(self) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate" or prop.appraisal_price:
            return
        if self.is_purchase_banned(prop):
            self.skip_banned_property()
            return
        prop.appraisal_mode = True
        self.comment = SHIMARISU_COMMENTS["appraisal_mode"]

    def submit_appraisal_offer(self, appraisal_price: int) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate" or prop.appraisal_price:
            return
        if self.is_purchase_banned(prop):
            self.skip_banned_property()
            return
        appraisal_price = max(500_000, round(appraisal_price / 100_000) * 100_000)
        fee = calc_brokerage_fee(appraisal_price) if prop.source == "broker" else 0
        if self.cash < appraisal_price + fee:
            self.comment = SHIMARISU_COMMENTS["cash_short"]
            self.add_log(UI_TEXT["log_cash_short"].format(amount=yen(appraisal_price + fee)))
            return

        accepted = random.random() < self.calculate_appraisal_acceptance(prop, appraisal_price)
        credit_delta, reason = self.appraisal_credit_change(prop, appraisal_price, accepted)
        prop.appraisal_price = appraisal_price
        prop.appraisal_reason = reason
        prop.appraisal_credit_delta = credit_delta
        prop.appraisal_mode = False
        self.reputation = max(0, min(100, self.reputation + credit_delta))

        if accepted:
            prop.appraisal_result = "accepted"
            self.add_log(UI_TEXT["appraisal_passed"])
            self.complete_purchase(appraisal_price, reason)
            if not self.pending_confirmation:
                self.pending_confirmation = "appraisal_accepted"
                self.comment = SHIMARISU_COMMENTS["appraisal_passed"]
            return

        prop.appraisal_result = "rejected"
        prop.status = "skipped"
        self.skipped_count += 1
        self.day += 2
        self.apply_operating_cost(2)
        self.comment = SHIMARISU_COMMENTS["appraisal_rejected"]
        self.add_log(UI_TEXT["appraisal_rejected"])
        self.pending_confirmation = "appraisal_rejected"

    def complete_purchase(self, purchase_price: int, appraisal_reason: str = "") -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate":
            return
        purchase_price = max(500_000, round(purchase_price / 100_000) * 100_000)
        fee = calc_brokerage_fee(purchase_price) if prop.source == "broker" else 0
        total = purchase_price + fee
        if self.cash < total:
            self.comment = SHIMARISU_COMMENTS["cash_short"]
            self.add_log(UI_TEXT["log_cash_short"].format(amount=yen(total)))
            return

        self.cash -= total
        prop.buy_price = purchase_price
        prop.brokerage_fee_buy = fee
        prop.status = "owned"
        prop.overpaid_purchase = self.is_overpaid_purchase(prop)
        if appraisal_reason and not prop.appraisal_reason:
            prop.appraisal_reason = appraisal_reason
        self.comment = SHIMARISU_COMMENTS["bought"]
        self.add_log(UI_TEXT["log_purchase"].format(amount=yen(total)))
        self.reveal_hidden_risks_after_purchase(prop)

    def reveal_hidden_risks_after_purchase(self, prop: PropertyCase) -> None:
        events: list[tuple[str, int, str, str]] = []

        def add_event(message: str, cost: int, key: str, source_key: str) -> None:
            events.append((message, cost, key, source_key))

        if not prop.investigated:
            if prop.rebuildable and random.random() < 0.22:
                prop.rebuildable = False
                if random.random() < 0.55:
                    prop.road_access_good = False
                    prop.rebuild_blocker_reason = "接道2m未満"
                    add_event(UI_TEXT["event_road_found"], 0, "road_access_good", "risk_investigate_short")
                else:
                    prop.rebuild_blocker_reason = random.choice(("市街化調整区域", "法令制限あり", "不明な制限あり"))
                    add_event(UI_TEXT["event_rebuild_found"], 0, "rebuildable", "risk_investigate_short")
            if not prop.illegal_building and random.random() < 0.12:
                prop.illegal_building = True
                add_event(UI_TEXT["event_illegal_found"], 0, "illegal_building", "risk_investigate_short")
            if not prop.flood_damage and random.random() < 0.16:
                prop.flood_damage = True
                add_event(UI_TEXT["event_flood_found"], 700_000, "flood_damage", "risk_investigate_short")
            if prop.rain_leak_level == "none" and random.random() < 0.18:
                prop.rain_leak_level = "suspected"
                add_event(UI_TEXT["event_rain_leak_found"], 600_000, "rain_leak_level", "risk_investigate_short")
            if prop.building_tilt_level == "none" and random.random() < 0.08:
                prop.building_tilt_level = "suspected"
                add_event(UI_TEXT["event_tilt_found"], 900_000, "building_tilt_level", "risk_investigate_short")

        if not prop.heard_from_seller:
            if not prop.incident_property and random.random() < 0.12:
                prop.incident_property = True
                add_event(UI_TEXT["event_incident_found"], 0, "incident_property", "risk_hearing_short")
            if not prop.termite_damage and random.random() < 0.22:
                prop.termite_level = "confirmed"
                prop.termite_damage = True
                add_event(UI_TEXT["event_termite_found"], 800_000, "termite_damage", "risk_hearing_short")
            if not prop.leftover_items and random.random() < 0.22:
                prop.leftover_level = "many"
                prop.leftover_items = True
                add_event(UI_TEXT["event_leftover_found"], 400_000, "leftover_items", "risk_hearing_short")
            if not prop.neighbor_trouble and random.random() < 0.18:
                prop.neighbor_trouble = True
                add_event(UI_TEXT["event_neighbor_found"], 0, "neighbor_trouble", "risk_hearing_short")
            if not prop.neighbor_garbage_house and random.random() < 0.08:
                prop.neighbor_garbage_house = True
                add_event(UI_TEXT["event_garbage_found"], 0, "neighbor_garbage_house", "risk_hearing_short")
            if not prop.nearby_antisocial_base and random.random() < 0.04:
                prop.nearby_antisocial_base = True
                add_event(UI_TEXT["event_nearby_antisocial_found"], 0, "nearby_antisocial_base", "risk_hearing_short")
            if prop.seller_antisocial_level == "none" and random.random() < 0.025:
                prop.seller_antisocial_level = "suspected"
                add_event(UI_TEXT["event_seller_antisocial_found"], 0, "seller_antisocial", "risk_hearing_short")
            if prop.seller_antisocial_level != "confirmed" and random.random() < 0.01:
                prop.seller_antisocial_level = "confirmed"
                prop.seller_antisocial = True
                add_event(UI_TEXT["event_seller_antisocial_found"], 0, "seller_antisocial", "risk_hearing_short")

        if prop.flood_damage and not prop.known_flags.get("flood_damage") and random.random() < 0.95:
            add_event(UI_TEXT["event_flood_found"], 700_000, "flood_damage", "risk_investigate_short")
        if prop.termite_damage and not prop.known_flags.get("termite_damage") and random.random() < 0.95:
            add_event(UI_TEXT["event_termite_found"], 800_000, "termite_damage", "risk_hearing_short")
        if prop.leftover_items and not prop.known_flags.get("leftover_items") and random.random() < 0.9:
            add_event(UI_TEXT["event_leftover_found"], 400_000, "leftover_items", "risk_hearing_short")
        if prop.incident_property and not prop.known_flags.get("incident_property") and random.random() < 0.9:
            add_event(UI_TEXT["event_incident_found"], 0, "incident_property", "risk_hearing_short")
        if prop.neighbor_trouble and not prop.known_flags.get("neighbor_trouble") and random.random() < 0.9:
            add_event(UI_TEXT["event_neighbor_found"], 0, "neighbor_trouble", "risk_hearing_short")
        if prop.rain_leak_level != "none" and not prop.known_flags.get("rain_leak_level") and random.random() < 0.9:
            cost = 600_000 if prop.rain_leak_level == "suspected" else 1_300_000
            add_event(UI_TEXT["event_rain_leak_found"], cost, "rain_leak_level", "risk_investigate_short")
        if prop.building_tilt_level != "none" and not prop.known_flags.get("building_tilt_level") and random.random() < 0.9:
            cost = 900_000 if prop.building_tilt_level == "suspected" else 1_800_000
            add_event(UI_TEXT["event_tilt_found"], cost, "building_tilt_level", "risk_investigate_short")
        if prop.neighbor_garbage_house and not prop.known_flags.get("neighbor_garbage_house") and random.random() < 0.85:
            add_event(UI_TEXT["event_garbage_found"], 0, "neighbor_garbage_house", "risk_hearing_short")
        if prop.nearby_antisocial_base and not prop.known_flags.get("nearby_antisocial_base") and random.random() < 0.9:
            add_event(UI_TEXT["event_nearby_antisocial_found"], 0, "nearby_antisocial_base", "risk_hearing_short")
        if prop.seller_antisocial_level != "none" and not prop.known_flags.get("seller_antisocial") and random.random() < 0.9:
            add_event(UI_TEXT["event_seller_antisocial_found"], 0, "seller_antisocial", "risk_hearing_short")
        if prop.illegal_building and not prop.known_flags.get("illegal_building") and random.random() < 0.95:
            add_event(UI_TEXT["event_illegal_found"], 0, "illegal_building", "risk_investigate_short")
        if not prop.rebuildable and not prop.known_flags.get("rebuildable") and random.random() < 0.95:
            add_event(UI_TEXT["event_rebuild_found"], 0, "rebuildable", "risk_investigate_short")
        if not prop.road_access_good and not prop.known_flags.get("road_access_good") and random.random() < 0.95:
            add_event(UI_TEXT["event_road_found"], 0, "road_access_good", "risk_investigate_short")

        if not events:
            return

        unique_events = []
        seen_keys: set[str] = set()
        for event in events:
            if event[2] in seen_keys:
                continue
            seen_keys.add(event[2])
            unique_events.append(event)

        total_extra = sum(cost for _, cost, _, _ in unique_events)
        for message, _, key, source_key in unique_events:
            prop.known_flags[key] = True
            self.add_log(UI_TEXT["log_purchase_event"].format(message=UI_TEXT[source_key].format(message=message)))
        if total_extra:
            prop.repair_cost_total += total_extra
            self.cash -= total_extra
            self.add_log(UI_TEXT["log_extra_cost"].format(amount=yen(total_extra)))
        prop.hidden_risk_found = True
        self.comment = SHIMARISU_COMMENTS["risk_found"]
        self.pending_confirmation = "risk"

    def skip_property(self) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate":
            return
        prop.status = "skipped"
        self.skipped_count += 1
        self.day += 2
        self.apply_operating_cost(2)
        if prop.source == "broker" and not self.has_known_high_risk(prop):
            self.reputation = max(0, self.reputation - 1)
            self.add_log(UI_TEXT["log_broker_skip_rep"])
        self.comment = SHIMARISU_COMMENTS["skip_good"] if self.has_known_high_risk(prop) else SHIMARISU_COMMENTS["skipped"]
        self.add_log(UI_TEXT["log_skip"].format(name=prop.name))
        self.add_log(UI_TEXT["log_skip_cost"])
        self.pending_confirmation = "skipped"

    def skip_banned_property(self) -> None:
        prop = self.current_property
        if not prop or prop.status != "candidate":
            return
        prop.status = "skipped"
        prop.banned_skip = True
        self.skipped_count += 1
        self.day += 1
        self.apply_operating_cost(1)
        self.comment = SHIMARISU_COMMENTS["risk_retreat"]
        self.add_log(UI_TEXT["purchase_banned_notice"])
        self.pending_confirmation = "skipped"

    def choose_renovation(self, option: dict[str, Any]) -> None:
        prop = self.current_property
        if not prop or prop.status != "owned" or prop.renovation_label_key:
            return
        cost = int(option["cost"])
        if self.cash < cost:
            self.comment = SHIMARISU_COMMENTS["cash_short"]
            self.add_log(UI_TEXT["log_cash_short_renovation"].format(amount=yen(cost)))
            return

        self.cash -= cost
        prop.renovation_budget = cost
        prop.renovation_days_left = int(option["days"])
        prop.renovation_label_key = str(option["label_key"])

        if cost == 0:
            self.comment = SHIMARISU_COMMENTS["as_is"]
            self.add_log(UI_TEXT["log_as_is"])
            return

        prop.status = "renovating"
        if cost >= 5_000_000:
            self.comment = SHIMARISU_COMMENTS["renovate_full"]
        elif cost >= 3_500_000:
            self.comment = SHIMARISU_COMMENTS["renovate_standard"]
        else:
            self.comment = SHIMARISU_COMMENTS["renovate_light"]
        self.add_log(
            UI_TEXT["log_renovation"].format(
                label=UI_TEXT[str(option["label_key"])],
                amount=yen(cost),
                days=option["days"],
            )
        )

    def is_sale_active(self, prop: PropertyCase) -> bool:
        return prop.status == "listed" or (prop.status == "renovating" and prop.listed_price > 0)

    def choose_list_price(self, option: dict[str, Any]) -> None:
        prop = self.current_property
        if not prop or prop.status not in {"owned", "renovating"} or not prop.renovation_label_key or prop.listed_price:
            return
        target = 2_500_000 if prop.renovation_budget > 0 else 1_500_000
        total_cost = self.calculate_total_cost(prop)
        listed_price = total_cost + target + int(option["profit_delta"])
        prop.target_profit = target
        prop.listed_price = max(1_000_000, round(listed_price / 100_000) * 100_000)
        prop.sale_strategy_key = str(option["key"])
        listing_during_renovation = prop.status == "renovating"
        if not listing_during_renovation:
            prop.status = "listed"
        prop.listed_days = 0
        prop.market_freshness = 100
        prop.last_sale_feedback = self.sale_feedback_reasons(prop)
        self.comment = SHIMARISU_COMMENTS["renovating_offer"] if listing_during_renovation else SHIMARISU_COMMENTS["listed"]
        self.add_log(
            UI_TEXT["log_listed"].format(
                label=UI_TEXT[str(option["label_key"])],
                amount=yen(prop.listed_price),
            )
        )
        if listing_during_renovation:
            self.add_log(UI_TEXT["renovation_listing"])

    def advance_day(self, allow_sale: bool = True) -> None:
        if self.screen_name != "main":
            return

        prop = self.current_property
        self.day += 1
        self.apply_operating_cost(1)

        if prop and prop.status in {"owned", "renovating", "listed"}:
            prop.days_held += 1
            if prop.days_held > 0 and prop.days_held % 30 == 0:
                fee = self.calculate_maintenance_fee(prop)
                self.cash -= fee
                prop.maintenance_cost_total += fee
                self.add_log(UI_TEXT["log_maintenance"].format(amount=yen(fee)))
                self.comment = SHIMARISU_COMMENTS["long_hold"]

        if prop and prop.status == "renovating":
            prop.renovation_days_left = max(0, prop.renovation_days_left - 1)
            if prop.renovation_days_left == 0:
                prop.status = "listed" if prop.listed_price else "owned"
                self.add_log(UI_TEXT["log_renovation_done"])
                self.comment = UI_TEXT["renovation_done_comment"]

        if prop and self.is_sale_active(prop):
            prop.listed_days += 1
            self.update_market_freshness(prop)
            prop.last_sale_feedback = self.sale_feedback_reasons(prop)

        if allow_sale and prop and self.is_sale_active(prop):
            self.try_sell_property(prop)

        if self.day > MAX_DAYS:
            self.finish_game()

    def finish_game(self) -> None:
        if not self.result_assessed:
            self.apply_unsold_inventory_penalty()
            self.result_assessed = True
        self.screen_name = "result"
        self.comment = self.result_comment()

    def apply_unsold_inventory_penalty(self) -> None:
        prop = self.current_property
        if not prop or prop.status not in {"owned", "renovating", "listed"}:
            return

        penalties: list[tuple[str, int]] = [(UI_TEXT["unsold_penalty_inventory"], -5)]
        if self.is_sale_active(prop) and prop.market_freshness < 50:
            penalties.append((UI_TEXT["unsold_penalty_stale"], -3))
        if self.unrealized_profit(prop) < 0:
            penalties.append((UI_TEXT["unsold_penalty_unrealized_loss"], -5))
        if self.cash < LOW_CASH_WARNING:
            penalties.append((UI_TEXT["unsold_penalty_low_cash"], -3))

        total_delta = sum(delta for _, delta in penalties)
        self.result_inventory_penalties = penalties
        self.reputation = max(0, self.reputation + total_delta)

    def try_sell_property(self, prop: PropertyCase) -> None:
        if self.pending_confirmation or prop.purchase_offer_price:
            return
        chance = self.calculate_sale_chance(prop)
        if random.random() >= chance:
            prop.last_sale_feedback = self.sale_feedback_reasons(prop)
            self.comment = SHIMARISU_COMMENTS["not_sold"]
            return

        self.create_purchase_offer(prop)

    def create_purchase_offer(self, prop: PropertyCase) -> None:
        self_buyer = random.random() < 0.62
        discount = self.purchase_offer_discount_for(prop)
        sold_price = max(500_000, prop.listed_price - discount)
        prop.purchase_offer_price = sold_price
        prop.purchase_offer_discount = discount
        prop.purchase_offer_self_buyer = self_buyer
        prop.purchase_offer_reason = self.purchase_offer_reason_for(prop, discount)
        self.pending_confirmation = "purchase_offer"
        self.comment = (
            SHIMARISU_COMMENTS["renovating_offer"]
            if prop.status == "renovating"
            else SHIMARISU_COMMENTS["offer_received"]
        )
        self.add_log(UI_TEXT["log_offer_received"].format(amount=yen(sold_price), discount=yen(discount)))

    def purchase_offer_discount_for(self, prop: PropertyCase) -> int:
        points = 0
        if prop.market_freshness < 85:
            points += 1
        if prop.listed_days >= 45:
            points += 1
        if prop.sale_strategy_key == "strong":
            points += 1
        if prop.overpaid_purchase or self.has_major_known_risk(prop):
            points += 1
        if prop.early_big_price_drop or prop.many_early_price_drops:
            points += 1
        if prop.nearby_antisocial_base or prop.building_tilt_level == "confirmed":
            points += 1
        if self.qualifies_under_construction_offer(prop):
            points = max(0, points - 1)
        if points <= 0:
            return 0
        if points == 1:
            return 300_000
        if points == 2:
            return 500_000
        if points == 3:
            return 800_000
        return 1_000_000

    def purchase_offer_reason_for(self, prop: PropertyCase, discount: int) -> str:
        if discount <= 0:
            return UI_TEXT["offer_reason_full"]
        if discount >= 1_000_000:
            return UI_TEXT["offer_reason_million"]
        if discount <= 300_000:
            return UI_TEXT["offer_reason_rounding"]
        if prop.market_freshness < 85 or prop.listed_days >= 45:
            return UI_TEXT["offer_reason_stale"]
        if self.has_major_known_risk(prop) or prop.nearby_antisocial_base or prop.building_tilt_level != "none":
            return UI_TEXT["offer_reason_risk"]
        return UI_TEXT["offer_reason_cash"]

    def accept_purchase_offer(self) -> None:
        prop = self.current_property
        if not prop or self.pending_confirmation != "purchase_offer" or not prop.purchase_offer_price:
            return
        sold_price = prop.purchase_offer_price
        discount = prop.purchase_offer_discount
        self_buyer = prop.purchase_offer_self_buyer
        total_cost = self.calculate_total_cost(prop)
        brokerage = calc_brokerage_fee(sold_price)
        outside_fee = 0 if self_buyer else calc_brokerage_fee(sold_price)
        profit = sold_price + brokerage - outside_fee - total_cost
        forced_credit = -3 if profit < 0 and discount > 0 else None
        forced_reason = UI_TEXT["offer_loss_credit"] if forced_credit is not None else None
        self.complete_sale(
            prop,
            sold_price,
            self_buyer,
            discount,
            forced_credit=forced_credit,
            forced_reason=forced_reason,
        )
        if profit >= 0 and discount > 0 and self.sale_target_met(prop, profit):
            prop.credit_reason = f"{prop.credit_reason}\n{UI_TEXT['offer_accepted_credit']}"
        self.add_log(UI_TEXT["log_offer_accepted"].format(amount=yen(sold_price)))

    def reject_purchase_offer(self) -> None:
        prop = self.current_property
        if not prop or self.pending_confirmation != "purchase_offer":
            return
        prop.purchase_offer_price = 0
        prop.purchase_offer_discount = 0
        prop.purchase_offer_reason = ""
        prop.market_freshness = max(35, prop.market_freshness - 3)
        prop.last_sale_feedback = self.sale_feedback_reasons(prop)
        self.pending_confirmation = None
        self.comment = SHIMARISU_COMMENTS["offer_reject"]
        self.add_log(UI_TEXT["log_offer_rejected"])

    def complete_sale(
        self,
        prop: PropertyCase,
        sold_price: int,
        self_buyer: bool,
        negotiation_discount: int,
        *,
        forced_credit: int | None = None,
        forced_reason: str | None = None,
    ) -> None:
        brokerage = calc_brokerage_fee(sold_price)
        outside_fee = 0 if self_buyer else calc_brokerage_fee(sold_price)
        proceeds = sold_price + brokerage - outside_fee
        total_cost = self.calculate_total_cost(prop)
        profit = proceeds - total_cost

        self.cash += proceeds
        sold_under_construction = prop.status == "renovating"
        prop.status = "sold"
        prop.sold_price = sold_price
        prop.sale_channel = "auto_customer" if self_buyer else "outside_customer"
        prop.brokerage_income = brokerage
        prop.outside_fee = outside_fee
        prop.net_proceeds = proceeds
        prop.sold_profit = profit
        prop.sale_negotiation_discount = negotiation_discount
        prop.sold_under_construction = sold_under_construction
        prop.purchase_offer_price = 0
        prop.purchase_offer_discount = 0
        prop.purchase_offer_reason = ""

        self.sold_count += 1
        if not prop.investigated:
            self.uninvestigated_sales += 1
        if not prop.heard_from_seller:
            self.unheard_sales += 1
        if profit < 0:
            self.loss_count += 1
        elif self.sale_target_missed(prop, profit):
            self.target_missed_sales += 1
        if profit > self.best_profit:
            self.best_profit = profit
            self.best_profit_name = prop.name

        prop.sale_result_comment = self.sale_result_comment_for(prop, profit)
        if forced_credit is None:
            credit_delta, credit_reason = self.calculate_credit_change(prop, profit)
        else:
            credit_delta, credit_reason = forced_credit, forced_reason or UI_TEXT["credit_loss"]
        prop.credit_delta = credit_delta
        prop.credit_reason = credit_reason
        self.reputation = max(0, min(100, self.reputation + credit_delta))

        self.comment = prop.sale_result_comment

        self.add_log(UI_TEXT["log_sold"].format(amount=yen(prop.sold_price), profit=yen(profit)))
        self.pending_confirmation = "sold"

    def can_loss_cut(self, prop: PropertyCase | None = None) -> bool:
        prop = prop or self.current_property
        if not prop or not self.is_sale_active(prop) or self.pending_confirmation:
            return False
        return prop.listed_days >= 60 or self.cash < LOW_CASH_WARNING or prop.market_freshness < 60

    def loss_cut_price_for(self, prop: PropertyCase) -> int:
        total_cost = self.calculate_total_cost(prop)
        raw_price = max(total_cost - 1_500_000, int(total_cost * 0.85))
        market_cap = int(self.risk_adjusted_market_value(prop) * 1.03)
        price = min(raw_price, market_cap)
        return max(500_000, round(price / 100_000) * 100_000)

    def loss_cut_property(self) -> None:
        prop = self.current_property
        if not self.can_loss_cut(prop):
            return

        assert prop is not None
        was_cash_short = self.cash < LOW_CASH_WARNING
        sold_price = self.loss_cut_price_for(prop)
        proceeds = sold_price
        total_cost = self.calculate_total_cost(prop)
        profit = proceeds - total_cost
        credit_delta = -1 if was_cash_short else -2

        self.cash += proceeds
        prop.status = "sold"
        prop.disposal_sale = True
        prop.sold_price = sold_price
        prop.sale_channel = "loss_cut_channel"
        prop.brokerage_income = 0
        prop.outside_fee = 0
        prop.net_proceeds = proceeds
        prop.sold_profit = profit
        prop.sale_negotiation_discount = max(0, prop.listed_price - sold_price)
        prop.credit_delta = credit_delta
        prop.credit_reason = UI_TEXT["credit_loss_cut"]

        self.sold_count += 1
        if not prop.investigated:
            self.uninvestigated_sales += 1
        if not prop.heard_from_seller:
            self.unheard_sales += 1
        if profit < 0:
            self.loss_count += 1
        elif self.sale_target_missed(prop, profit):
            self.target_missed_sales += 1
        if profit > self.best_profit:
            self.best_profit = profit
            self.best_profit_name = prop.name

        prop.sale_result_comment = self.sale_result_comment_for(prop, profit)
        self.reputation = max(0, min(100, self.reputation + credit_delta))
        self.comment = SHIMARISU_COMMENTS["loss_cut"]
        self.add_log(UI_TEXT["log_loss_cut"].format(amount=yen(sold_price), profit=yen(profit)))
        self.pending_confirmation = "sold"

    def calculate_maintenance_fee(self, prop: PropertyCase) -> int:
        fee = BASE_MAINTENANCE_FEE
        if prop.age >= 35:
            fee += 20_000
        if prop.flood_damage:
            fee += 20_000
        if prop.termite_damage:
            fee += 30_000
        if prop.illegal_building:
            fee += 10_000
        if prop.rain_leak_level == "suspected":
            fee += 20_000
        elif prop.rain_leak_level == "confirmed":
            fee += 45_000
        if prop.building_tilt_level == "suspected":
            fee += 25_000
        elif prop.building_tilt_level == "confirmed":
            fee += 60_000
        return fee

    def calculate_total_cost(self, prop: PropertyCase) -> int:
        return (
            prop.buy_price
            + prop.brokerage_fee_buy
            + prop.renovation_budget
            + prop.repair_cost_total
            + prop.maintenance_cost_total
        )

    def target_profit_for(self, prop: PropertyCase) -> int | None:
        if prop.status in {"candidate", "skipped"}:
            return None
        if not prop.renovation_label_key:
            return None
        return 2_500_000 if prop.renovation_budget > 0 else 1_500_000

    def projected_profit_for(self, prop: PropertyCase, price: int | None = None) -> int | None:
        if prop.status == "sold":
            return prop.sold_profit
        if price is None:
            if prop.listed_price:
                price = prop.listed_price
            elif prop.renovation_label_key:
                price = self.preview_sale_price(prop, SALE_OPTIONS[1])
            else:
                return None
        return price + calc_brokerage_fee(price) - self.calculate_total_cost(prop)

    def sale_target_met(self, prop: PropertyCase, profit: int) -> bool:
        target = self.target_profit_for(prop)
        return target is not None and profit >= target

    def sale_target_missed(self, prop: PropertyCase, profit: int) -> bool:
        target = self.target_profit_for(prop)
        return target is not None and 0 <= profit < target

    def sale_result_comment_for(self, prop: PropertyCase, profit: int) -> str:
        if profit < 0:
            return UI_TEXT["sale_result_reform_loss"]
        if self.sale_target_met(prop, profit):
            if prop.renovation_budget > 0:
                return UI_TEXT["sale_result_target_clear"]
            return UI_TEXT["sale_result_as_is_clear"]
        if prop.renovation_budget > 0:
            return f"{UI_TEXT['sale_result_reform_thin']}\n{UI_TEXT['sale_result_black_target_missed']}"
        return f"{UI_TEXT['sale_result_as_is_thin']}\n{UI_TEXT['sale_result_black_target_missed']}"

    def buyer_finance_status(self, prop: PropertyCase) -> str:
        if prop.illegal_building:
            return "cash_only"
        if not prop.rebuildable:
            return "cash_center" if not prop.road_access_good else "loan_hard"
        if not prop.road_access_good:
            return "cash_center"
        return "loan_available"

    def is_purchase_banned(self, prop: PropertyCase) -> bool:
        return prop.seller_antisocial_level == "confirmed" and prop.known_flags.get("seller_antisocial", False)

    def has_wide_parking_mismatch(self, prop: PropertyCase) -> bool:
        return prop.floor_plan in {"4LDK", "5DK", "5LDK"} and prop.parking_count <= 1

    def qualifies_under_construction_offer(self, prop: PropertyCase) -> bool:
        if prop.status != "renovating" or not prop.listed_price:
            return False
        projected = self.projected_profit_for(prop) or 0
        target = self.target_profit_for(prop) or 0
        _, recommended_high, _ = self.recommended_appraisal_range(prop)
        return (
            prop.buy_price <= recommended_high
            and projected >= target
            and prop.station_walk_minutes <= 15
            and prop.parking_count >= 2
            and not self.has_major_known_risk(prop)
            and prop.sale_strategy_key != "strong"
        )

    def appraisal_visible(self, prop: PropertyCase) -> bool:
        return prop.investigated or prop.heard_from_seller

    def known_risk_score(self, prop: PropertyCase) -> int:
        score = 0
        if prop.known_flags.get("rebuildable") and not prop.rebuildable:
            score += 3
        if prop.known_flags.get("road_access_good") and not prop.road_access_good:
            score += 3
        if prop.known_flags.get("illegal_building") and prop.illegal_building:
            score += 3
        if prop.known_flags.get("incident_property") and prop.incident_property:
            score += 2
        if prop.known_flags.get("termite_damage") and prop.termite_damage:
            score += 3 if prop.termite_level == "confirmed" else 2
        if prop.known_flags.get("flood_damage") and prop.flood_damage:
            score += 2
        if prop.known_flags.get("leftover_items") and prop.leftover_items:
            score += 2 if prop.leftover_level == "many" else 1
        if prop.known_flags.get("rain_leak_level") and prop.rain_leak_level == "suspected":
            score += 2
        elif prop.known_flags.get("rain_leak_level") and prop.rain_leak_level == "confirmed":
            score += 3
        if prop.known_flags.get("building_tilt_level") and prop.building_tilt_level == "suspected":
            score += 3
        elif prop.known_flags.get("building_tilt_level") and prop.building_tilt_level == "confirmed":
            score += 5
        if prop.known_flags.get("neighbor_garbage_house") and prop.neighbor_garbage_house:
            score += 2
        if prop.known_flags.get("nearby_antisocial_base") and prop.nearby_antisocial_base:
            score += 5
        if prop.known_flags.get("seller_antisocial") and prop.seller_antisocial_level == "suspected":
            score += 3
        if self.is_purchase_banned(prop):
            score += 8
        if prop.parking_count == 0:
            score += 1
        if self.has_wide_parking_mismatch(prop):
            score += 1
        if prop.station_walk_minutes > 20:
            score += 1
        return score

    def has_major_known_risk(self, prop: PropertyCase) -> bool:
        return self.known_risk_score(prop) >= 3

    def recommended_appraisal_range(self, prop: PropertyCase) -> tuple[int, int, str]:
        expected = prop.expected_sale_price
        target_profit = 2_500_000
        standard_renovation = 3_500_000
        repair_estimate = 0
        reason = UI_TEXT["appraisal_ok_reason"]

        if self.is_purchase_banned(prop):
            return 0, 0, UI_TEXT["purchase_banned_notice"]

        if prop.known_flags.get("illegal_building") and prop.illegal_building:
            low, high = int(expected * 0.40), int(expected * 0.55)
            reason = UI_TEXT["feedback_illegal"]
        elif prop.known_flags.get("nearby_antisocial_base") and prop.nearby_antisocial_base:
            low, high = int(expected * 0.35), int(expected * 0.50)
            reason = UI_TEXT["feedback_nearby_antisocial"]
        elif (
            prop.known_flags.get("rebuildable")
            and not prop.rebuildable
            or prop.known_flags.get("road_access_good")
            and not prop.road_access_good
        ):
            low, high = int(expected * 0.45), int(expected * 0.60)
            reason = UI_TEXT["feedback_rebuild"] if not prop.rebuildable else UI_TEXT["feedback_road"]
        elif prop.known_flags.get("incident_property") and prop.incident_property:
            low, high = int(expected * 0.55), int(expected * 0.70)
            reason = UI_TEXT["feedback_incident"]
        else:
            if prop.known_flags.get("termite_damage") and prop.termite_damage:
                repair_estimate += 800_000
                reason = UI_TEXT["feedback_termite"]
            if prop.known_flags.get("flood_damage") and prop.flood_damage:
                repair_estimate += 700_000
                reason = UI_TEXT["feedback_flood"]
            if prop.known_flags.get("rain_leak_level") and prop.rain_leak_level != "none":
                repair_estimate += 600_000 if prop.rain_leak_level == "suspected" else 1_300_000
                reason = UI_TEXT["feedback_rain_leak"]
            if prop.known_flags.get("building_tilt_level") and prop.building_tilt_level != "none":
                repair_estimate += 900_000 if prop.building_tilt_level == "suspected" else 1_800_000
                reason = UI_TEXT["feedback_tilt"]
            center = expected - target_profit - standard_renovation - repair_estimate
            if prop.station_walk_minutes > 30:
                center -= 700_000
                reason = UI_TEXT["feedback_station_far"]
            elif prop.station_walk_minutes > 20:
                center -= 400_000
                reason = UI_TEXT["feedback_station_far"]
            if prop.floor_plan == "2LDK":
                center -= 300_000
                reason = UI_TEXT["feedback_floor_2ldk"]
            elif prop.floor_plan in {"5DK", "4DK"}:
                center -= 200_000
                reason = UI_TEXT["feedback_floor_old"]
            if self.has_wide_parking_mismatch(prop):
                center -= 600_000 if prop.parking_count == 0 else 300_000
                reason = UI_TEXT["feedback_parking_floor_mismatch"]
            low, high = center - 500_000, center + 500_000

        low = max(1_000_000, round(low / 100_000) * 100_000)
        high = max(low + 500_000, round(high / 100_000) * 100_000)
        return low, high, reason

    def calculate_appraisal_acceptance(self, prop: PropertyCase, appraisal_price: int) -> float:
        asking = max(1, prop.seller_asking_price or prop.buy_price)
        ratio = appraisal_price / asking
        if ratio >= 0.98:
            chance = 0.95
        elif ratio >= 0.90:
            chance = 0.80
        elif ratio >= 0.80:
            chance = 0.60
        elif ratio >= 0.70:
            chance = 0.40
        elif ratio >= 0.60:
            chance = 0.25
        else:
            chance = 0.10

        if prop.known_flags.get("rebuildable") and not prop.rebuildable:
            chance += 0.15
        if prop.known_flags.get("road_access_good") and not prop.road_access_good:
            chance += 0.15
        if prop.known_flags.get("illegal_building") and prop.illegal_building:
            chance += 0.15
        if prop.known_flags.get("incident_property") and prop.incident_property:
            chance += 0.10
        if prop.known_flags.get("termite_damage") and prop.termite_damage:
            chance += 0.10
        if prop.known_flags.get("flood_damage") and prop.flood_damage:
            chance += 0.10
        if prop.known_flags.get("leftover_items") and prop.leftover_items:
            chance += 0.05
        if prop.known_flags.get("rain_leak_level") and prop.rain_leak_level != "none":
            chance += 0.08
        if prop.known_flags.get("building_tilt_level") and prop.building_tilt_level != "none":
            chance += 0.12
        if prop.known_flags.get("neighbor_garbage_house") and prop.neighbor_garbage_house:
            chance += 0.06
        if prop.known_flags.get("nearby_antisocial_base") and prop.nearby_antisocial_base:
            chance += 0.18
        if prop.station_walk_minutes > 20:
            chance += 0.05
        if prop.parking_count == 0:
            chance += 0.05
        if self.known_risk_score(prop) <= 1 and ratio < 0.80:
            chance -= 0.10
        if self.reputation >= 70:
            chance += 0.05
        elif self.reputation < 40:
            chance -= 0.10
        return clamp(chance, 0.05, 0.98)

    def appraisal_credit_change(self, prop: PropertyCase, appraisal_price: int, accepted: bool) -> tuple[int, str]:
        asking = max(1, prop.seller_asking_price or prop.buy_price)
        ratio = appraisal_price / asking
        risk_score = self.known_risk_score(prop)
        recommended_low, recommended_high, reason = self.recommended_appraisal_range(prop)
        if risk_score <= 1:
            if ratio < 0.60:
                return -6, UI_TEXT["appraisal_lowball_reason"]
            if ratio < 0.70:
                return -4, UI_TEXT["appraisal_lowball_reason"]
            if ratio < 0.80:
                return -2, UI_TEXT["appraisal_lowball_reason"]
        if accepted and prop.investigated and prop.heard_from_seller and recommended_low <= appraisal_price <= recommended_high:
            return 1, reason
        if accepted and risk_score >= 3 and ratio >= 0.90:
            return 0, UI_TEXT["appraisal_high_risk_reason"]
        if not accepted and ratio < 0.80:
            return -1, UI_TEXT["appraisal_lowball_reason"]
        return 0, reason

    def is_overpaid_purchase(self, prop: PropertyCase) -> bool:
        _, recommended_high, _ = self.recommended_appraisal_range(prop)
        asking = max(1, prop.seller_asking_price or prop.buy_price)
        ratio = prop.buy_price / asking
        return prop.buy_price > recommended_high or (self.has_major_known_risk(prop) and ratio >= 0.90)

    def appraisal_options(self, prop: PropertyCase) -> list[tuple[str, int]]:
        asking = prop.seller_asking_price or prop.buy_price
        options: list[tuple[str, int]] = [
            (UI_TEXT["offer_at_asking"], asking),
            (UI_TEXT["offer_minus_30"], asking - 300_000),
            (UI_TEXT["offer_minus_50"], asking - 500_000),
            (UI_TEXT["offer_minus_80"], asking - 800_000),
            (UI_TEXT["offer_minus_100"], asking - 1_000_000),
        ]

        deduped: list[tuple[str, int]] = []
        seen: set[int] = set()
        for label, price in options:
            rounded = max(500_000, round(price / 100_000) * 100_000)
            if rounded in seen:
                continue
            seen.add(rounded)
            deduped.append((label, rounded))
        return deduped[:5]

    def calculate_credit_change(self, prop: PropertyCase, profit: int) -> tuple[int, str]:
        if profit < 0:
            return -3, UI_TEXT["credit_loss"]

        target_met = self.sale_target_met(prop, profit)
        target_missed = self.sale_target_missed(prop, profit)
        discount_taken = prop.sale_negotiation_discount > 0

        if prop.sold_under_construction:
            if discount_taken and target_missed:
                return -1, UI_TEXT["credit_under_construction_discount"]
            if discount_taken and target_met:
                base = 2 if prop.investigated and prop.heard_from_seller else 1
                return base, UI_TEXT["credit_under_construction_discount"]
            if target_met and prop.investigated and prop.heard_from_seller:
                return 4, UI_TEXT["sale_result_target_clear"]
            if target_missed:
                return 0, UI_TEXT["credit_target_missed"]

        if target_missed:
            base = -1 if prop.renovation_budget > 0 else 0
            return base, UI_TEXT["credit_target_missed"]

        reason = UI_TEXT["credit_full_good"]
        if prop.hidden_risk_found and not (prop.investigated and prop.heard_from_seller):
            base = -2
            reason = UI_TEXT["credit_hidden_risk"]
        elif prop.overpaid_purchase:
            base = 0 if target_met else -1
            reason = UI_TEXT["appraisal_high_risk_reason"]
        elif prop.investigated and prop.heard_from_seller:
            base = 3
            reason = UI_TEXT["credit_full_good"]
        elif target_met:
            base = 0
            reason = UI_TEXT["credit_no_check_good"]
        else:
            base = -1
            reason = UI_TEXT["credit_no_check_good"]

        if target_met and prop.days_held <= 30:
            base += 2
            reason = f"{reason}\n{UI_TEXT['quick_exit']}"
        elif target_met and prop.days_held <= 60:
            base += 1
            reason = f"{reason}\n{UI_TEXT['quick_exit']}"
        elif prop.days_held > 120:
            base -= 1
            reason = f"{reason}\n{UI_TEXT['slow_exit']}"
        return max(-5, min(6, base)), reason

    def exit_label_for(self, prop: PropertyCase) -> str:
        if prop.renovation_label_key == "as_is":
            return UI_TEXT["old_house_land_exit"]
        return UI_TEXT[prop.renovation_label_key]

    def sale_negotiation_discount_for(self, prop: PropertyCase) -> int:
        points = 0
        if prop.price_cut_count >= 2:
            points += 1
        if prop.listed_days >= 60:
            points += 1
        if prop.sale_strategy_key == "strong":
            points += 1
        if not prop.rebuildable or prop.illegal_building:
            points += 1
        if prop.termite_damage or prop.flood_damage or prop.neighbor_trouble:
            points += 1
        if prop.parking_count == 0 or prop.age >= 40:
            points += 1
        if prop.station_walk_minutes > 20 or prop.floor_plan == "2LDK":
            points += 1
        if prop.overpaid_purchase:
            points += 1
        if not prop.investigated:
            points += 1
        if not prop.heard_from_seller:
            points += 1
        if points <= 1:
            return 0
        rate = min(0.08, 0.015 * points)
        return min(1_000_000, round((prop.listed_price * rate) / 100_000) * 100_000)

    def update_market_freshness(self, prop: PropertyCase) -> None:
        if prop.listed_days > 90:
            freshness = 65
        elif prop.listed_days > 60:
            freshness = 80
        elif prop.listed_days > 30:
            freshness = 90
        else:
            freshness = 100
        freshness += min(12, prop.price_cut_count * 3)
        if prop.early_big_price_drop:
            freshness -= 10
        if prop.many_early_price_drops:
            freshness -= 15
        prop.market_freshness = max(35, min(100, freshness))

    def has_known_high_risk(self, prop: PropertyCase) -> bool:
        checks = (
            (not prop.rebuildable, "rebuildable"),
            (not prop.road_access_good, "road_access_good"),
            (prop.illegal_building, "illegal_building"),
            (prop.incident_property, "incident_property"),
            (prop.termite_damage, "termite_damage"),
            (prop.rain_leak_level != "none", "rain_leak_level"),
            (prop.building_tilt_level != "none", "building_tilt_level"),
            (prop.nearby_antisocial_base, "nearby_antisocial_base"),
            (prop.seller_antisocial_level != "none", "seller_antisocial"),
        )
        return any(is_risky and prop.known_flags.get(key, False) for is_risky, key in checks)

    def is_price_floor(self, prop: PropertyCase) -> bool:
        floor_price = self.calculate_total_cost(prop) + 500_000
        return self.is_sale_active(prop) and prop.listed_price - 500_000 < floor_price

    def discount_listed_price(self, amount: int) -> None:
        prop = self.current_property
        if not prop or not self.is_sale_active(prop) or self.pending_confirmation:
            return
        floor_price = self.calculate_total_cost(prop) + 500_000
        next_price = prop.listed_price - amount
        if next_price < floor_price:
            self.comment = f"{SHIMARISU_COMMENTS['price_down_blocked']}\n{UI_TEXT['price_down_blocked_extra']}"
            prop.last_sale_feedback = self.sale_feedback_reasons(prop)
            self.add_log(f"{UI_TEXT['price_down_blocked']} {UI_TEXT['price_down_blocked_extra']}")
            return
        prop.listed_price = round(next_price / 100_000) * 100_000
        prop.price_cut_count += 1
        prop.price_drop_total += amount
        prop.market_freshness = min(100, prop.market_freshness + (6 if amount >= 1_000_000 else 3))
        if prop.listed_days <= 7 and amount >= 1_000_000:
            prop.early_big_price_drop = True
            prop.market_freshness = max(35, prop.market_freshness - 10)
            self.comment = SHIMARISU_COMMENTS["lowball_market"]
        elif prop.listed_days <= 14 and prop.price_cut_count >= 2:
            prop.many_early_price_drops = True
            prop.market_freshness = max(35, prop.market_freshness - 15)
            self.comment = SHIMARISU_COMMENTS["lowball_market"]
        else:
            self.comment = SHIMARISU_COMMENTS["price_down"]
        prop.last_sale_feedback = self.sale_feedback_reasons(prop)
        self.add_log(UI_TEXT["log_price_down"].format(amount=yen(amount)))

    def holding_value(self, prop: PropertyCase | None = None) -> int:
        prop = prop or self.current_property
        if not prop or prop.status not in {"owned", "renovating", "listed"}:
            return 0
        if prop.listed_price:
            return prop.listed_price
        return prop.expected_sale_price

    def total_assets_estimate(self) -> int:
        return self.cash + self.holding_value()

    def unrealized_profit(self, prop: PropertyCase | None = None) -> int:
        prop = prop or self.current_property
        if not prop or prop.status not in {"owned", "renovating", "listed"}:
            return 0
        return self.holding_value(prop) - self.calculate_total_cost(prop)

    def sale_feedback_reasons(self, prop: PropertyCase) -> list[str]:
        reasons: list[str] = []
        if self.is_price_floor(prop):
            reasons.append(UI_TEXT["feedback_price_floor"])
        if prop.early_big_price_drop:
            reasons.append(UI_TEXT["feedback_early_big_drop"])
        if prop.many_early_price_drops:
            reasons.append(UI_TEXT["feedback_many_price_drops"])
        if prop.listed_days >= 60 and prop.price_cut_count == 0:
            reasons.append(UI_TEXT["feedback_need_price_review"])
        if prop.overpaid_purchase:
            reasons.append(UI_TEXT["feedback_purchase_heavy"])
            if self.has_major_known_risk(prop):
                reasons.append(UI_TEXT["feedback_risk_purchase_heavy"])
        market = self.risk_adjusted_market_value(prop)
        if prop.listed_price > market * 1.04:
            reasons.append(UI_TEXT["feedback_price_high"])
        if prop.sale_strategy_key == "strong":
            reasons.append(UI_TEXT["feedback_strategy_strong"])
        if not prop.rebuildable:
            reasons.append(UI_TEXT["feedback_rebuild"])
        if not prop.road_access_good:
            reasons.append(UI_TEXT["feedback_road"])
        if prop.illegal_building:
            reasons.append(UI_TEXT["feedback_illegal"])
        if prop.incident_property:
            reasons.append(UI_TEXT["feedback_incident"])
        if prop.flood_damage:
            reasons.append(UI_TEXT["feedback_flood"])
        if prop.termite_damage:
            reasons.append(UI_TEXT["feedback_termite"])
        if prop.rain_leak_level != "none":
            reasons.append(UI_TEXT["feedback_rain_leak"])
        if prop.building_tilt_level != "none":
            reasons.append(UI_TEXT["feedback_tilt"])
        if prop.parking_count == 0:
            reasons.append(UI_TEXT["feedback_parking"])
        elif prop.parking_count == 1:
            reasons.append(UI_TEXT["feedback_suburban_one_parking"])
        if self.has_wide_parking_mismatch(prop):
            reasons.append(
                UI_TEXT["feedback_parking_floor_mismatch"]
                if prop.parking_count == 0
                else UI_TEXT["feedback_family_parking"]
            )
        if prop.station_walk_minutes > 20:
            reasons.append(UI_TEXT["feedback_station_far"])
        if prop.floor_plan == "2LDK":
            reasons.append(UI_TEXT["feedback_floor_2ldk"])
        elif prop.floor_plan in {"5DK", "5LDK", "4DK"} and prop.renovation_budget < 3_500_000:
            reasons.append(UI_TEXT["feedback_floor_old"])
        if prop.age >= 40:
            reasons.append(UI_TEXT["feedback_old"])
        if prop.days_held >= 60:
            reasons.append(UI_TEXT["feedback_long"])
        if prop.market_freshness < 90:
            reasons.append(UI_TEXT["feedback_market_stale"])
        if not prop.investigated:
            reasons.append(UI_TEXT["feedback_under_investigated"])
        if not prop.heard_from_seller:
            reasons.append(UI_TEXT["feedback_under_heard"])
        if self.reputation < 45:
            reasons.append(UI_TEXT["feedback_reputation"])
        if self.buyer_finance_status(prop) in {"cash_center", "cash_only"}:
            reasons.append(UI_TEXT["feedback_cash_center"])
        if prop.neighbor_garbage_house:
            reasons.append(UI_TEXT["feedback_neighbor_garbage"])
        if prop.nearby_antisocial_base:
            reasons.append(UI_TEXT["feedback_nearby_antisocial"])
        return reasons[:4] or [UI_TEXT["feedback_normal"]]

    def candidate_attention_points(self, prop: PropertyCase) -> list[str]:
        points: list[str] = []
        if not prop.investigated:
            points.append(UI_TEXT["investigate_hint"])
        if not prop.heard_from_seller:
            points.append(UI_TEXT["hearing_hint"])
        asking = prop.seller_asking_price or prop.buy_price
        rough_profit = prop.expected_sale_price - asking
        if rough_profit < 2_000_000:
            points.append(UI_TEXT["skip_reason_margin"])
        if prop.parking_count == 0:
            points.append(UI_TEXT["feedback_parking"])
        if self.has_wide_parking_mismatch(prop):
            points.append(UI_TEXT["feedback_parking_floor_mismatch"])
        if prop.station_walk_minutes > 20:
            points.append(UI_TEXT["feedback_station_far"])
        if prop.floor_plan == "2LDK":
            points.append(UI_TEXT["feedback_floor_2ldk"])
        if prop.age >= 40:
            points.append(UI_TEXT["feedback_old"])
        if prop.seller_antisocial_level != "none" and prop.known_flags.get("seller_antisocial"):
            points.insert(0, f"{UI_TEXT['seller_antisocial']}: {self.seller_antisocial_text(prop)}")
        if prop.rain_leak_level != "none" and prop.known_flags.get("rain_leak_level"):
            points.append(UI_TEXT["feedback_rain_leak"])
        if prop.building_tilt_level != "none" and prop.known_flags.get("building_tilt_level"):
            points.append(UI_TEXT["feedback_tilt"])
        if prop.termite_level != "none" and prop.known_flags.get("termite_damage"):
            points.append(f"{UI_TEXT['termite']}: {self.termite_level_text(prop.termite_level)}")
        if prop.leftover_level != "none" and prop.known_flags.get("leftover_items"):
            points.append(f"{UI_TEXT['leftover']}: {self.leftover_level_text(prop.leftover_level)}")
        if prop.nearby_antisocial_base and prop.known_flags.get("nearby_antisocial_base"):
            points.append(UI_TEXT["feedback_nearby_antisocial"])
        if prop.neighbor_garbage_house and prop.known_flags.get("neighbor_garbage_house"):
            points.append(UI_TEXT["feedback_neighbor_garbage"])
        if self.is_purchase_banned(prop):
            points.insert(0, UI_TEXT["purchase_banned_notice"])
        return points[:5]

    def skip_reason_for(self, prop: PropertyCase) -> str:
        if prop.banned_skip or self.is_purchase_banned(prop):
            return UI_TEXT["purchase_banned"]
        if self.has_known_high_risk(prop):
            return UI_TEXT["skip_reason_risk"]
        asking = prop.seller_asking_price or prop.buy_price
        rough_profit = prop.expected_sale_price - asking
        if rough_profit < 2_000_000:
            return UI_TEXT["skip_reason_margin"]
        if self.cash < asking + (calc_brokerage_fee(asking) if prop.source == "broker" else 0):
            return UI_TEXT["skip_reason_cash"]
        return UI_TEXT["skip_reason_optional"]

    def renovation_value_bonus(self, prop: PropertyCase) -> int:
        if prop.renovation_budget <= 0:
            return 0
        if prop.age < 15:
            multiplier = 0.48
        elif prop.age < 35:
            multiplier = 0.72
        else:
            multiplier = 0.95
        return int(prop.renovation_budget * multiplier)

    def risk_adjusted_market_value(self, prop: PropertyCase) -> int:
        value = prop.expected_sale_price + self.renovation_value_bonus(prop)
        if prop.parking_count == 0:
            value -= 700_000
        elif prop.parking_count >= 2:
            value += 700_000
        if prop.corner_lot:
            value += 400_000
        if not prop.rebuildable:
            value -= 2_200_000
        if not prop.road_access_good:
            value -= 900_000
        if prop.incident_property:
            value -= 1_300_000
        if prop.flood_damage:
            value -= 700_000
        if prop.termite_damage:
            value -= 800_000
        if prop.illegal_building:
            value -= 1_600_000
        if prop.rain_leak_level == "suspected":
            value -= 600_000
        elif prop.rain_leak_level == "confirmed":
            value -= 1_300_000
        if prop.building_tilt_level == "suspected":
            value -= 1_000_000
        elif prop.building_tilt_level == "confirmed":
            value -= 2_300_000
        if prop.leftover_items and prop.renovation_budget == 0:
            value -= 650_000
        if prop.neighbor_trouble:
            value -= 500_000
        if prop.neighbor_garbage_house:
            value -= 700_000
        if prop.nearby_antisocial_base:
            value -= 2_700_000
        if prop.station_walk_minutes <= 10:
            value += 600_000
        elif prop.station_walk_minutes > 30:
            value -= 1_200_000
        elif prop.station_walk_minutes > 20:
            value -= 700_000
        if prop.parking_count >= 2 and prop.station_walk_minutes > 20:
            value += 350_000
        if prop.floor_plan in {"3LDK", "4LDK"}:
            value += 400_000
        elif prop.floor_plan == "2LDK":
            value -= 500_000
        elif prop.floor_plan in {"5DK", "5LDK", "4DK"} and prop.renovation_budget < 3_500_000:
            value -= 300_000
        if self.has_wide_parking_mismatch(prop):
            value -= 900_000 if prop.parking_count == 0 else 400_000
        return max(1_000_000, value)

    def calculate_sale_chance(self, prop: PropertyCase) -> float:
        chance = 0.085
        if prop.sale_strategy_key == "early":
            chance += 0.075
        elif prop.sale_strategy_key == "fair":
            chance += 0.025
        elif prop.sale_strategy_key == "strong":
            chance -= 0.04

        if prop.parking_count == 0:
            chance -= 0.025
        elif prop.parking_count >= 2:
            chance += 0.025
        else:
            chance += 0.01
        if prop.station_walk_minutes <= 10:
            chance += 0.025
        elif prop.station_walk_minutes > 30:
            chance -= 0.055
        elif prop.station_walk_minutes > 20:
            chance -= 0.035
        if prop.parking_count >= 2 and prop.station_walk_minutes > 20:
            chance += 0.018
        if prop.floor_plan in {"3LDK", "4LDK"}:
            chance += 0.018
        elif prop.floor_plan == "2LDK":
            chance -= 0.02
        elif prop.floor_plan in {"5DK", "5LDK", "4DK"} and prop.renovation_budget >= 3_500_000:
            chance += 0.015
        elif prop.floor_plan in {"5DK", "5LDK", "4DK"}:
            chance -= 0.012
        if self.has_wide_parking_mismatch(prop):
            chance -= 0.045 if prop.parking_count == 0 else 0.025
        if prop.corner_lot:
            chance += 0.015
        if prop.age < 15:
            chance += 0.02
        elif prop.age >= 45:
            chance -= 0.04
        elif prop.age >= 35:
            chance -= 0.025
        if prop.renovation_budget >= 3_500_000:
            chance += 0.04
        elif prop.renovation_budget >= 1_500_000:
            chance += 0.025
        elif prop.age >= 35:
            chance -= 0.015

        if not prop.rebuildable:
            chance -= 0.055
        if not prop.road_access_good:
            chance -= 0.035
        if prop.incident_property:
            chance -= 0.04
        if prop.flood_damage:
            chance -= 0.025
        if prop.termite_damage:
            chance -= 0.03
        if prop.illegal_building:
            chance -= 0.06
        if prop.rain_leak_level == "suspected":
            chance -= 0.025
        elif prop.rain_leak_level == "confirmed":
            chance -= 0.05
        if prop.building_tilt_level == "suspected":
            chance -= 0.055
        elif prop.building_tilt_level == "confirmed":
            chance -= 0.10
        if prop.leftover_items and prop.renovation_budget == 0:
            chance -= 0.02
        if prop.neighbor_trouble:
            chance -= 0.02
        if prop.neighbor_garbage_house:
            chance -= 0.03
        if prop.nearby_antisocial_base:
            chance -= 0.12
        if prop.overpaid_purchase:
            chance -= 0.035
        if prop.early_big_price_drop:
            chance -= 0.025
        if prop.many_early_price_drops:
            chance -= 0.035
        if prop.listed_days >= 60 and prop.price_cut_count == 0:
            chance -= 0.03
        if self.qualifies_under_construction_offer(prop):
            chance += 0.055
        chance += min(0.04, prop.price_cut_count * 0.012)
        if not prop.investigated and not prop.heard_from_seller:
            chance -= 0.30
        elif not prop.investigated:
            chance -= 0.15
        elif not prop.heard_from_seller:
            chance -= 0.10
        chance -= (100 - prop.market_freshness) * 0.0012
        if prop.sale_strategy_key == "strong" and prop.market_freshness < 80:
            chance -= 0.03

        if prop.days_held >= 90:
            chance -= 0.03
        elif prop.days_held >= 60:
            chance -= 0.015
        chance += (self.reputation - 50) * 0.001

        market = self.risk_adjusted_market_value(prop)
        ratio = prop.listed_price / max(1, market)
        if ratio > 1.0:
            chance -= min(0.07, (ratio - 1.0) * 0.25)
        else:
            chance += min(0.05, (1.0 - ratio) * 0.18)

        if prop.illegal_building:
            chance = min(chance, 0.075)
        if prop.nearby_antisocial_base:
            chance = min(chance, 0.055)
        return clamp(chance, 0.012, 0.32)

    def result_comment(self) -> str:
        if self.unsold_count():
            prop = self.current_property
            if prop and self.unrealized_profit(prop) < 0:
                return UI_TEXT["result_unsold_loss"]
            if self.cash < LOW_CASH_WARNING:
                return UI_TEXT["result_unsold_cash_short"]
            return UI_TEXT["result_unsold_asset_plus"]
        if self.target_missed_sales:
            return UI_TEXT["trust_comment_cash_recovered"]
        gain = self.cash - INITIAL_CASH
        if gain >= 4_000_000:
            return SHIMARISU_COMMENTS["result_good"]
        if gain >= 0:
            return SHIMARISU_COMMENTS["result_even"]
        return SHIMARISU_COMMENTS["result_bad"]

    def draw(self) -> None:
        self.buttons = []
        if self.screen_name == "title":
            self.draw_title()
        elif self.screen_name == "result":
            self.draw_result()
        else:
            self.draw_main()

    def draw_title(self) -> None:
        self.screen.fill(COLORS["bg"])
        if self.village:
            bg = pygame.transform.smoothscale(self.village, (WINDOW_WIDTH, WINDOW_HEIGHT))
            bg.set_alpha(86)
            self.screen.blit(bg, (0, 0))
        overlay = pygame.Surface((WINDOW_WIDTH, WINDOW_HEIGHT), pygame.SRCALPHA)
        overlay.fill((255, 255, 255, 180))
        self.screen.blit(overlay, (0, 0))

        title_font = self.get_font(46, True)
        sub_font = self.get_font(18)
        title_surf = title_font.render(UI_TEXT["title"], True, COLORS["text"])
        self.screen.blit(title_surf, title_surf.get_rect(center=(WINDOW_WIDTH // 2, 104)))
        sub_surf = sub_font.render(UI_TEXT["subtitle"], True, COLORS["muted"])
        self.screen.blit(sub_surf, sub_surf.get_rect(center=(WINDOW_WIDTH // 2, 152)))

        image_rect = pygame.Rect(0, 0, 308, 308)
        image_rect.center = (WINDOW_WIDTH // 2, 330)
        self.draw_shimarisu_image(self.shimarisu_title, image_rect, large=True)
        self.draw_wrapped_text(
            self.comment,
            pygame.Rect(WINDOW_WIDTH // 2 - 220, 500, 440, 54),
            self.get_font(15),
            COLORS["muted"],
            center=True,
        )
        self.add_button(pygame.Rect(WINDOW_WIDTH // 2 - 96, 584, 192, 54), UI_TEXT["start"], "start")
        self.draw_buttons()

    def draw_main(self) -> None:
        self.screen.fill(COLORS["bg"])
        self.draw_header()
        self.draw_left_panel()
        self.draw_center_panel()
        self.draw_right_panel()
        self.draw_bottom_buttons()

    def draw_result(self) -> None:
        self.screen.fill(COLORS["bg"])
        panel = pygame.Rect(160, 36, 860, 728)
        self.draw_panel(panel, COLORS["panel"])
        self.draw_text(UI_TEXT["result"], (panel.x + 32, panel.y + 28), 34, True)

        rows = [
            (UI_TEXT["cash_on_hand"], yen(self.cash)),
            (UI_TEXT["holding_value"], yen(self.holding_value())),
            (UI_TEXT["total_assets"], yen(self.total_assets_estimate())),
            (UI_TEXT["asset_gain"], yen(self.total_assets_estimate() - INITIAL_CASH)),
            (UI_TEXT["operating_cost_total"], yen(self.operating_cost_total)),
            (UI_TEXT["sold_count"], f"{self.sold_count}件"),
            (UI_TEXT["skipped_count"], f"{self.skipped_count}件"),
            (UI_TEXT["unsold_count"], f"{self.unsold_count()}件"),
            (UI_TEXT["loss_count"], f"{self.loss_count}件"),
            (UI_TEXT["target_missed_sales"], f"{self.target_missed_sales}件"),
            (UI_TEXT["uninvestigated_sales"], f"{self.uninvestigated_sales}件"),
            (UI_TEXT["unheard_sales"], f"{self.unheard_sales}件"),
        ]
        if self.best_profit_name:
            rows.extend(
                [
                    (UI_TEXT["best_profit"], self.best_profit_name),
                    (UI_TEXT["profit"], yen(self.best_profit)),
                ]
            )
        else:
            rows.append((UI_TEXT["best_profit"], UI_TEXT["best_profit_none"]))
        rows.append((UI_TEXT["reputation"], f"{self.reputation}"))
        row_y = panel.y + 104
        row_h = 27
        split = (len(rows) + 1) // 2
        for index, (label, value) in enumerate(rows):
            column = 0 if index < split else 1
            row_index = index if column == 0 else index - split
            x = panel.x + 42 + column * 400
            y = row_y + row_index * row_h
            self.draw_key_value(label, value, pygame.Rect(x, y, 350, 24))

        self.draw_inventory_evaluation(panel, panel.y + 318)
        if self.shimarisu:
            self.draw_shimarisu_image(self.shimarisu, pygame.Rect(panel.x + 42, panel.y + 486, 96, 96))
        self.draw_text(UI_TEXT["trust_evaluation"], (panel.x + 156, panel.y + 486), 17, True, COLORS["green_dark"])
        self.draw_wrapped_text(
            self.trust_evaluation_comment(),
            pygame.Rect(panel.x + 156, panel.y + 514, 620, 62),
            self.get_font(16),
            COLORS["text"],
        )
        self.draw_wrapped_text(
            self.result_detail_comment(),
            pygame.Rect(panel.x + 156, panel.y + 586, 620, 42),
            self.get_font(13),
            COLORS["muted"],
        )
        self.add_button(pygame.Rect(panel.x + 278, panel.y + 660, 136, 48), UI_TEXT["restart"], "restart")
        self.add_button(pygame.Rect(panel.x + 446, panel.y + 660, 136, 48), UI_TEXT["quit"], "quit")
        self.draw_buttons()

    def best_profit_text(self) -> str:
        if self.best_profit_name:
            return yen(self.best_profit)
        return UI_TEXT["best_profit_none"]

    def unsold_count(self) -> int:
        prop = self.current_property
        return 1 if prop and prop.status in {"owned", "renovating", "listed"} else 0

    def result_detail_comment(self) -> str:
        if self.cash < 0:
            return f"{self.comment}\n{UI_TEXT['cash_short_result']}"
        if self.cash < LOW_CASH_WARNING:
            return f"{self.comment}\n{UI_TEXT['cash_warning']}"
        return self.comment

    def trust_evaluation_comment(self) -> str:
        if self.unsold_count():
            if self.cash < LOW_CASH_WARNING or self.unrealized_profit() < 0:
                return UI_TEXT["trust_comment_bad"]
            return UI_TEXT["result_unsold_asset_plus"]
        if self.target_missed_sales >= 2:
            return UI_TEXT["trust_comment_target_missed"]
        if self.target_missed_sales == 1:
            return UI_TEXT["trust_comment_thin_exit"]
        if self.reputation >= 58 and self.loss_count == 0 and self.uninvestigated_sales == 0 and self.unheard_sales == 0:
            return UI_TEXT["trust_comment_good"]
        if self.total_assets_estimate() >= INITIAL_CASH and (self.uninvestigated_sales or self.unheard_sales):
            return UI_TEXT["trust_comment_mixed"]
        if self.reputation < 45 or self.total_assets_estimate() < INITIAL_CASH or self.loss_count:
            return UI_TEXT["trust_comment_bad"]
        if self.unsold_count():
            return SHIMARISU_COMMENTS["result_even"]
        return UI_TEXT["trust_comment_good"]

    def draw_inventory_evaluation(self, panel: pygame.Rect, y: int) -> None:
        box = pygame.Rect(panel.x + 42, y, panel.width - 84, 142)
        pygame.draw.rect(self.screen, COLORS["panel_alt"], box, border_radius=8)
        pygame.draw.rect(self.screen, COLORS["line"], box, 1, border_radius=8)
        self.draw_text(UI_TEXT["unsold_evaluation"], (box.x + 14, box.y + 12), 16, True, COLORS["muted"])
        summary = UI_TEXT["unsold_has_inventory"] if self.result_inventory_penalties else UI_TEXT["unsold_clear"]
        self.draw_wrapped_text(summary, pygame.Rect(box.x + 14, box.y + 40, box.width - 28, 34), self.get_font(13), COLORS["text"])
        self.draw_text(UI_TEXT["unsold_credit_change"], (box.x + 14, box.y + 82), 14, True, COLORS["muted"])
        if not self.result_inventory_penalties:
            self.draw_text(UI_TEXT["none"], (box.x + 128, box.y + 82), 14, False, COLORS["muted"])
            return
        x = box.x + 128
        y_line = box.y + 82
        for label, delta in self.result_inventory_penalties[:4]:
            self.draw_text(f"{label} {self.signed_plain(delta)}", (x, y_line), 13, False, COLORS["red"])
            x += 150
            if x > box.right - 140:
                x = box.x + 128
                y_line += 24

    def draw_header(self) -> None:
        header = pygame.Rect(24, 18, 1132, 64)
        self.draw_panel(header, COLORS["panel"])
        self.draw_text(UI_TEXT["title"], (header.x + 22, header.y + 17), 24, True)
        stats = [
            (UI_TEXT["day"], f"{min(self.day, MAX_DAYS)}日目 / {MAX_DAYS}日"),
            (UI_TEXT["cash"], yen(self.cash)),
            (UI_TEXT["reputation"], str(self.reputation)),
            (UI_TEXT["daily_operating_cost"], f"{yen(DAILY_OPERATING_COST)}/日"),
        ]
        x = header.right - 724
        for label, value in stats:
            self.draw_stat_chip(label, value, pygame.Rect(x, header.y + 14, 172, 36))
            x += 184

    def draw_left_panel(self) -> None:
        panel = pygame.Rect(24, 98, 270, 606)
        self.draw_panel(panel, COLORS["panel"])
        self.draw_text(UI_TEXT["shimarisu_name"], (panel.x + 18, panel.y + 18), 20, True)
        self.draw_shimarisu_image(
            self.shimarisu,
            pygame.Rect(panel.x + 42, panel.y + 54, 186, 166),
            animated=True,
        )
        self.draw_wrapped_text(
            self.comment,
            pygame.Rect(panel.x + 18, panel.y + 236, 234, 66),
            self.get_font(15),
            COLORS["text"],
        )
        self.draw_office_area(pygame.Rect(panel.x + 18, panel.y + 320, 234, 142))
        self.draw_text(UI_TEXT["memo"], (panel.x + 18, panel.y + 486), 15, True, COLORS["muted"])
        log_y = panel.y + 514
        for log in self.message_log[:2]:
            self.draw_wrapped_text(
                f"・{log}",
                pygame.Rect(panel.x + 18, log_y, 234, 28),
                self.get_font(12),
                COLORS["muted"],
            )
            log_y += 29

    def draw_center_panel(self) -> None:
        panel = pygame.Rect(314, 98, 470, 606)
        self.draw_panel(panel, COLORS["panel"])
        prop = self.current_property
        if not prop or prop.status == "candidate":
            title = UI_TEXT["current_case"]
        elif prop.status == "skipped":
            title = UI_TEXT["skipped_done"]
        else:
            title = UI_TEXT["owned_property"]
        self.draw_text(title, (panel.x + 20, panel.y + 18), 21, True)

        if not prop:
            self.draw_wrapped_text(
                UI_TEXT["no_property"],
                pygame.Rect(panel.x + 24, panel.y + 90, 390, 80),
                self.get_font(18),
                COLORS["muted"],
                center=True,
            )
            return

        self.draw_text(prop.name, (panel.x + 20, panel.y + 56), 23, True, COLORS["green_dark"])
        self.draw_status_badge(prop, pygame.Rect(panel.right - 130, panel.y + 58, 104, 28))

        rows = [
            (UI_TEXT["area"], prop.area),
            (UI_TEXT["age"], f"築{prop.age}年"),
            (UI_TEXT["floor_plan"], prop.floor_plan),
            (UI_TEXT["station_distance"], self.station_distance_text(prop)),
            (UI_TEXT["seller_asking_price"], yen(prop.seller_asking_price or prop.buy_price)),
            (UI_TEXT["expected_sale"], yen(prop.expected_sale_price)),
            (UI_TEXT["source"], UI_TEXT[prop.source]),
            (UI_TEXT["parking"], f"{prop.parking_count}台"),
            (UI_TEXT["corner_lot"], self.yes_no(prop.corner_lot)),
        ]
        if prop.status not in {"candidate", "skipped"}:
            rows.insert(5, (UI_TEXT["purchase_price"], yen(prop.buy_price)))
        y = panel.y + 110
        for label, value in rows:
            self.draw_key_value(label, value, pygame.Rect(panel.x + 24, y, 390, 24))
            y += 28

        self.draw_text(UI_TEXT["known_info"], (panel.x + 24, y + 8), 16, True, COLORS["muted"])
        y += 38
        for label, value, known in self.visible_risk_rows(prop):
            if y + 30 > panel.bottom - 8:
                break
            if value == "" and known:
                self.draw_text(label, (panel.x + 24, y + 4), 13, True, COLORS["green_dark"])
                y += 24
                continue
            text_value = value if known else UI_TEXT["not_checked"]
            color = COLORS["text"] if known else COLORS["muted"]
            self.draw_key_value(label, text_value, pygame.Rect(panel.x + 24, y, 390, 26), value_color=color)
            y += 30

        if prop.status == "sold" and y + 76 < panel.bottom:
            banner = pygame.Rect(panel.x + 24, y + 12, 422, 62)
            is_loss = prop.sold_profit < 0
            banner_bg = (255, 235, 232) if is_loss else COLORS["soft_green"]
            banner_line = COLORS["red"] if is_loss else COLORS["green"]
            banner_color = COLORS["red"] if is_loss else COLORS["green_dark"]
            profit_label = UI_TEXT["confirmed_loss"] if is_loss else UI_TEXT["confirmed_profit"]
            result_title = UI_TEXT["loss_cut_result"] if prop.disposal_sale else UI_TEXT["sold_banner"]
            pygame.draw.rect(self.screen, banner_bg, banner, border_radius=8)
            pygame.draw.rect(self.screen, banner_line, banner, 1, border_radius=8)
            self.draw_text(result_title, (banner.x + 16, banner.y + 10), 17, True, banner_color)
            profit_surf = self.get_font(20, True).render(
                f"{profit_label}: {yen(prop.sold_profit)}",
                True,
                banner_color,
            )
            self.screen.blit(profit_surf, (banner.x + 16, banner.y + 34))
        elif self.pending_confirmation == "risk" and y + 64 < panel.bottom:
            banner = pygame.Rect(panel.x + 24, y + 12, 422, 54)
            pygame.draw.rect(self.screen, (255, 238, 218), banner, border_radius=8)
            pygame.draw.rect(self.screen, COLORS["orange"], banner, 1, border_radius=8)
            self.draw_text(UI_TEXT["risk_notice"], (banner.x + 16, banner.y + 10), 17, True, COLORS["orange"])
            self.draw_wrapped_text(
                self.comment,
                pygame.Rect(banner.x + 16, banner.y + 31, banner.width - 32, 20),
                self.get_font(12),
                COLORS["muted"],
            )

    def draw_right_panel(self) -> None:
        panel = pygame.Rect(804, 98, 352, 606)
        self.draw_panel(panel, COLORS["panel"])
        prop = self.current_property
        if not prop:
            return
        self.draw_text(self.right_panel_title(prop), (panel.x + 18, panel.y + 18), 20, True)

        if prop.status == "candidate":
            self.draw_candidate_right_panel(panel, prop)
            return
        if prop.status == "skipped":
            self.draw_skipped_right_panel(panel, prop)
            return
        if self.is_sale_active(prop):
            self.draw_listed_right_panel(panel, prop)
            return

        y = panel.y + 58
        cost_rows = [
            (UI_TEXT["brokerage_buy"], yen(prop.brokerage_fee_buy)),
            (UI_TEXT["renovation_cost"], yen(prop.renovation_budget)),
            (UI_TEXT["repair_cost"], yen(prop.repair_cost_total)),
            (UI_TEXT["maintenance"], yen(prop.maintenance_cost_total)),
            (UI_TEXT["total_cost"], yen(self.calculate_total_cost(prop))),
        ]
        for label, value in cost_rows:
            self.draw_key_value(label, value, pygame.Rect(panel.x + 18, y, 316, 24), small=True)
            y += 28

        if self.pending_confirmation == "appraisal_accepted" and prop.appraisal_result == "accepted":
            y += 10
            box = pygame.Rect(panel.x + 18, y, 316, 166)
            pygame.draw.rect(self.screen, COLORS["soft_green"], box, border_radius=8)
            pygame.draw.rect(self.screen, COLORS["green"], box, 1, border_radius=8)
            self.draw_text(UI_TEXT["appraisal_passed"], (box.x + 14, box.y + 10), 15, True, COLORS["green_dark"])
            rows = [
                (UI_TEXT["seller_asking_price"], yen(prop.seller_asking_price or prop.buy_price)),
                (UI_TEXT["appraisal_price"], yen(prop.appraisal_price)),
                (UI_TEXT["purchase_price"], yen(prop.buy_price)),
                (UI_TEXT["credit_delta"], self.signed_plain(prop.appraisal_credit_delta)),
            ]
            row_y = box.y + 42
            for label, value in rows:
                color = COLORS["green_dark"] if label == UI_TEXT["credit_delta"] and prop.appraisal_credit_delta >= 0 else COLORS["text"]
                if label == UI_TEXT["credit_delta"] and prop.appraisal_credit_delta < 0:
                    color = COLORS["red"]
                self.draw_key_value(label, value, pygame.Rect(box.x + 14, row_y, box.width - 28, 20), small=True, value_color=color)
                row_y += 23
            self.draw_wrapped_text(prop.appraisal_reason, pygame.Rect(box.x + 14, row_y + 4, box.width - 28, 42), self.get_font(12), COLORS["muted"])
            return

        if prop.status == "sold":
            y += 10
            profit_rect = pygame.Rect(panel.x + 18, y, 316, 76)
            is_loss = prop.sold_profit < 0
            box_bg = (255, 235, 232) if is_loss else COLORS["soft_green"]
            box_line = COLORS["red"] if is_loss else COLORS["green"]
            profit_label = UI_TEXT["confirmed_loss"] if is_loss else UI_TEXT["confirmed_profit"]
            profit_color = COLORS["red"] if is_loss else COLORS["green_dark"]
            result_title = UI_TEXT["loss_cut_result"] if prop.disposal_sale else UI_TEXT["sold_banner"]
            pygame.draw.rect(self.screen, box_bg, profit_rect, border_radius=8)
            pygame.draw.rect(self.screen, box_line, profit_rect, 1, border_radius=8)
            self.draw_text(result_title, (profit_rect.x + 14, profit_rect.y + 10), 16, True, profit_color)
            self.draw_text(profit_label, (profit_rect.x + 14, profit_rect.y + 42), 13, False, COLORS["muted"])
            profit_surf = self.get_font(19, True).render(yen(prop.sold_profit), True, profit_color)
            self.screen.blit(profit_surf, profit_surf.get_rect(midright=(profit_rect.right - 14, profit_rect.y + 52)))
            y += 94

            target = self.target_profit_for(prop)
            sold_rows = [
                (UI_TEXT["sale_price"], yen(prop.sold_price)),
                (UI_TEXT["total_cost"], yen(self.calculate_total_cost(prop))),
                (UI_TEXT["sale_negotiation"], f"-{yen(prop.sale_negotiation_discount)}"),
                (UI_TEXT["brokerage_income"], yen(prop.brokerage_income)),
                (UI_TEXT["outside_fee"], yen(prop.outside_fee)),
                (UI_TEXT["net_proceeds"], yen(prop.net_proceeds)),
                (UI_TEXT["sale_channel"], UI_TEXT[prop.sale_channel]),
            ]
            if target is not None:
                sold_rows.append((UI_TEXT["target_profit"], yen(target)))
                sold_rows.append((UI_TEXT["target_gap"], self.signed_yen(prop.sold_profit - target)))
            sold_rows.append((UI_TEXT["credit_delta"], self.signed_plain(prop.credit_delta)))
            self.draw_text(UI_TEXT["sold_result"], (panel.x + 18, y), 16, True, COLORS["muted"])
            y += 30
            for label, value in sold_rows:
                value_color = COLORS["green_dark"] if label == UI_TEXT["net_proceeds"] else COLORS["text"]
                if label == UI_TEXT["target_gap"]:
                    value_color = COLORS["green_dark"] if prop.sold_profit >= (target or 0) else COLORS["red"]
                if label == UI_TEXT["credit_delta"]:
                    value_color = COLORS["green_dark"] if prop.credit_delta >= 0 else COLORS["red"]
                self.draw_key_value(
                    label,
                    value,
                    pygame.Rect(panel.x + 18, y, 316, 22),
                    small=True,
                    value_color=value_color,
                )
                y += 25
            if y + 66 < panel.bottom:
                comment_lines = prop.sale_result_comment or prop.credit_reason
                if prop.credit_reason and prop.credit_reason not in comment_lines:
                    comment_lines = f"{comment_lines}\n{prop.credit_reason}"
                self.draw_wrapped_text(comment_lines, pygame.Rect(panel.x + 18, y + 6, 316, 62), self.get_font(11), COLORS["muted"])
            return

        y += 10
        self.draw_profit_targets(panel, y)
        y += 104

        if prop.status in {"owned", "renovating", "listed", "sold"}:
            hold_rows = [
                (UI_TEXT["days_held"], f"{prop.days_held}日"),
                (UI_TEXT["status"], UI_TEXT[prop.status]),
                (UI_TEXT["buyer_type"], UI_TEXT[self.buyer_finance_status(prop)]),
            ]
            if prop.renovation_label_key:
                label = UI_TEXT["exit_method"] if prop.renovation_label_key == "as_is" else UI_TEXT["renovate"]
                hold_rows.append((label, self.exit_label_for(prop)))
            if prop.status == "renovating":
                hold_rows.append((UI_TEXT["remaining_work"], f"{prop.renovation_days_left}日"))
            if prop.listed_price:
                hold_rows.append((UI_TEXT["listed_price"], yen(prop.listed_price)))
                hold_rows.append((UI_TEXT["sale_chance"], f"{self.calculate_sale_chance(prop) * 100:.1f}%"))
            for label, value in hold_rows:
                self.draw_key_value(label, value, pygame.Rect(panel.x + 18, y, 316, 22), small=True)
                y += 25
            if prop.status == "listed":
                y += 8
                self.draw_sale_feedback(panel, y)
                return

        if prop.renovation_label_key and prop.status in {"owned", "renovating"} and not prop.listed_price:
            y += 10
            self.draw_text(UI_TEXT["select_price"], (panel.x + 18, y), 15, True, COLORS["muted"])
            y += 30
            for option in SALE_OPTIONS:
                price = self.preview_sale_price(prop, option)
                self.draw_key_value(
                    UI_TEXT[str(option["label_key"])],
                    yen(price),
                    pygame.Rect(panel.x + 18, y, 316, 24),
                    small=True,
                )
                y += 23
                self.draw_key_value(
                    UI_TEXT["projected_profit"],
                    yen(self.projected_profit_for(prop, price) or 0),
                    pygame.Rect(panel.x + 38, y, 296, 22),
                    small=True,
                    value_color=COLORS["green_dark"],
                )
                y += 25
            market = self.risk_adjusted_market_value(prop)
            self.draw_key_value(
                UI_TEXT["market_feel"],
                yen(market),
                pygame.Rect(panel.x + 18, y + 6, 316, 24),
                small=True,
                value_color=COLORS["green_dark"],
            )
        else:
            y += 10
            self.draw_cash_memo(panel, y)

    def right_panel_title(self, prop: PropertyCase) -> str:
        if self.pending_confirmation == "purchase_offer":
            return UI_TEXT["offer_received"]
        if self.pending_confirmation == "appraisal_accepted" and prop.appraisal_result == "accepted":
            return UI_TEXT["appraisal_offer_title"]
        if prop.status == "candidate":
            if prop.appraisal_mode:
                return UI_TEXT["appraisal_offer_title"]
            return UI_TEXT["case_memo"]
        if prop.status == "skipped":
            return UI_TEXT["skip_result"]
        if prop.status == "renovating":
            if prop.listed_price:
                return UI_TEXT["renovating_listed"]
            return UI_TEXT["renovating"]
        if prop.status == "listed":
            return UI_TEXT["listed"]
        if prop.status == "sold":
            return UI_TEXT["sold_result"]
        return UI_TEXT["cost_profit"]

    def draw_candidate_right_panel(self, panel: pygame.Rect, prop: PropertyCase) -> None:
        y = panel.y + 58
        asking = prop.seller_asking_price or prop.buy_price
        rough_profit = prop.expected_sale_price - asking
        rows = [
            (UI_TEXT["seller_asking_price"], yen(asking)),
            (UI_TEXT["expected_sale"], yen(prop.expected_sale_price)),
            (UI_TEXT["rough_profit"], yen(rough_profit)),
            (UI_TEXT["investigation_state"], UI_TEXT["survey_done"] if prop.investigated else UI_TEXT["not_checked"]),
            (UI_TEXT["hearing_state"], UI_TEXT["hearing_done"] if prop.heard_from_seller else UI_TEXT["not_checked"]),
            (UI_TEXT["daily_operating_cost"], f"{yen(DAILY_OPERATING_COST)} / 日"),
        ]
        for label, value in rows:
            self.draw_key_value(label, value, pygame.Rect(panel.x + 18, y, 316, 24), small=True)
            y += 28

        y += 4
        self.draw_text(UI_TEXT["recommended_appraisal"], (panel.x + 18, y), 15, True, COLORS["muted"])
        y += 26
        if self.is_purchase_banned(prop):
            self.draw_wrapped_text(UI_TEXT["purchase_banned_notice"], pygame.Rect(panel.x + 18, y, 316, 54), self.get_font(12), COLORS["red"])
            y += 58
        elif self.appraisal_visible(prop):
            low, high, reason = self.recommended_appraisal_range(prop)
            self.draw_key_value(
                UI_TEXT["recommended_appraisal"],
                f"{yen(low)} - {yen(high)}",
                pygame.Rect(panel.x + 18, y, 316, 22),
                small=True,
                value_color=COLORS["green_dark"],
            )
            y += 26
            self.draw_wrapped_text(reason, pygame.Rect(panel.x + 18, y, 316, 38), self.get_font(12), COLORS["muted"])
            y += 42
        else:
            self.draw_wrapped_text(UI_TEXT["recommended_after_check"], pygame.Rect(panel.x + 18, y, 316, 28), self.get_font(12), COLORS["muted"])
            y += 34

        y += 8
        self.draw_text(UI_TEXT["attention_points"], (panel.x + 18, y), 15, True, COLORS["muted"])
        y += 26
        points = self.candidate_attention_points(prop)
        for point in points[:5]:
            self.draw_wrapped_text(
                f"・{point}",
                pygame.Rect(panel.x + 18, y, 316, 22),
                self.get_font(11),
                COLORS["muted"],
            )
            y += 22

        y += 8
        if not prop.investigated:
            self.draw_wrapped_text(UI_TEXT["appraisal_warning"], pygame.Rect(panel.x + 18, y, 316, 38), self.get_font(11), COLORS["red"])
            y += 40
        elif not prop.heard_from_seller:
            self.draw_wrapped_text(UI_TEXT["appraisal_warning_hearing"], pygame.Rect(panel.x + 18, y, 316, 38), self.get_font(11), COLORS["red"])

    def draw_skipped_right_panel(self, panel: pygame.Rect, prop: PropertyCase) -> None:
        y = panel.y + 58
        message = UI_TEXT["appraisal_rejected"] if prop.appraisal_result == "rejected" else UI_TEXT["skip_message"]
        if prop.banned_skip:
            message = UI_TEXT["purchase_banned_notice"]
        self.draw_wrapped_text(message, pygame.Rect(panel.x + 18, y, 316, 34), self.get_font(15), COLORS["text"])
        y += 48
        rep_delta = self.signed_plain(prop.appraisal_credit_delta) if prop.appraisal_result == "rejected" else (
            "-1" if prop.source == "broker" and not self.has_known_high_risk(prop) else "なし"
        )
        if prop.banned_skip:
            rep_delta = "なし"
        rows = [
            (UI_TEXT["seller_asking_price"], yen(prop.seller_asking_price or prop.buy_price)),
            (UI_TEXT["expected_sale"], yen(prop.expected_sale_price)),
            (UI_TEXT["skip_reason"], self.skip_reason_for(prop)),
            (UI_TEXT["elapsed_days"], "+1日" if prop.banned_skip else "+2日"),
            (UI_TEXT["credit_delta"], rep_delta),
        ]
        if prop.appraisal_price:
            rows.insert(2, (UI_TEXT["appraisal_price"], yen(prop.appraisal_price)))
        for label, value in rows:
            self.draw_key_value(label, value, pygame.Rect(panel.x + 18, y, 316, 24), small=True)
            y += 29
        y += 12
        self.draw_text(UI_TEXT["cash_memo"], (panel.x + 18, y), 15, True, COLORS["muted"])
        y += 28
        detail = prop.appraisal_reason if prop.appraisal_result == "rejected" else self.comment
        self.draw_wrapped_text(detail, pygame.Rect(panel.x + 18, y, 316, 74), self.get_font(13), COLORS["muted"])

    def draw_listed_right_panel(self, panel: pygame.Rect, prop: PropertyCase) -> None:
        y = panel.y + 58
        projected = self.projected_profit_for(prop) or 0
        rows = [
            (UI_TEXT["listed_price"], yen(prop.listed_price)),
            (UI_TEXT["projected_profit"], yen(projected)),
            (UI_TEXT["sale_chance"], f"{self.calculate_sale_chance(prop) * 100:.1f}%"),
            (UI_TEXT["sales_days"], f"{prop.listed_days}日"),
            (UI_TEXT["market_freshness"], f"{prop.market_freshness}"),
        ]
        for label, value in rows:
            self.draw_key_value(label, value, pygame.Rect(panel.x + 18, y, 316, 23), small=True)
            y += 28

        if self.pending_confirmation == "purchase_offer" and prop.purchase_offer_price:
            y += 8
            box = pygame.Rect(panel.x + 18, y, 316, 176)
            pygame.draw.rect(self.screen, COLORS["soft_green"], box, border_radius=8)
            pygame.draw.rect(self.screen, COLORS["green"], box, 1, border_radius=8)
            title = UI_TEXT["under_construction_offer"] if prop.status == "renovating" else UI_TEXT["offer_received"]
            self.draw_text(title, (box.x + 14, box.y + 10), 15, True, COLORS["green_dark"])
            offer_rows = [
                (UI_TEXT["listed_price"], yen(prop.listed_price), COLORS["text"]),
                (UI_TEXT["offer_price"], yen(prop.purchase_offer_price), COLORS["green_dark"]),
                (UI_TEXT["offer_discount"], f"-{yen(prop.purchase_offer_discount)}", COLORS["red"] if prop.purchase_offer_discount else COLORS["text"]),
            ]
            row_y = box.y + 42
            for label, value, color in offer_rows:
                self.draw_key_value(label, value, pygame.Rect(box.x + 14, row_y, box.width - 28, 20), small=True, value_color=color)
                row_y += 24
            self.draw_text(UI_TEXT["offer_reason"], (box.x + 14, row_y + 4), 12, True, COLORS["muted"])
            self.draw_wrapped_text(
                prop.purchase_offer_reason,
                pygame.Rect(box.x + 14, row_y + 24, box.width - 28, 44),
                self.get_font(12),
                COLORS["muted"],
            )
            return

        if self.can_loss_cut(prop):
            self.draw_key_value(
                UI_TEXT["loss_cut_price"],
                yen(self.loss_cut_price_for(prop)),
                pygame.Rect(panel.x + 18, y, 316, 23),
                small=True,
                value_color=COLORS["red"],
            )
            y += 28

        y += 8
        self.draw_sale_feedback(panel, y)
        y += 128
        self.draw_text(UI_TEXT["cash_assets"], (panel.x + 18, y), 15, True, COLORS["muted"])
        y += 28
        asset_rows = [
            (UI_TEXT["cash_on_hand"], yen(self.cash)),
            (UI_TEXT["holding_value"], yen(self.holding_value(prop))),
            (UI_TEXT["total_assets"], yen(self.total_assets_estimate())),
            (UI_TEXT["unrealized_profit"], self.signed_yen(self.unrealized_profit(prop))),
        ]
        for label, value in asset_rows:
            self.draw_key_value(label, value, pygame.Rect(panel.x + 18, y, 316, 21), small=True)
            y += 24
        if self.cash < LOW_CASH_WARNING:
            self.draw_wrapped_text(UI_TEXT["cash_warning"], pygame.Rect(panel.x + 18, y + 4, 316, 38), self.get_font(12), COLORS["red"])

    def draw_profit_targets(self, panel: pygame.Rect, y: int) -> None:
        prop = self.current_property
        if not prop:
            return
        box = pygame.Rect(panel.x + 18, y, 316, 92)
        pygame.draw.rect(self.screen, COLORS["panel_alt"], box, border_radius=8)
        pygame.draw.rect(self.screen, COLORS["line"], box, 1, border_radius=8)

        target = self.target_profit_for(prop)
        projected = self.projected_profit_for(prop)
        if target is None:
            rows = [
                (UI_TEXT["target_profit"], UI_TEXT["after_strategy"], COLORS["muted"]),
                (UI_TEXT["projected_profit"], UI_TEXT["after_price"], COLORS["muted"]),
                (UI_TEXT["target_gap"], UI_TEXT["after_price"], COLORS["muted"]),
            ]
        else:
            projected_value = projected or 0
            gap = projected_value - target
            rows = [
                (UI_TEXT["target_profit"], yen(target), COLORS["text"]),
                (UI_TEXT["projected_profit"], yen(projected_value), COLORS["green_dark"] if projected_value >= target else COLORS["red"]),
                (UI_TEXT["target_gap"], self.signed_yen(gap), COLORS["green_dark"] if gap >= 0 else COLORS["red"]),
            ]

        row_y = box.y + 10
        for label, value, color in rows:
            self.draw_key_value(
                label,
                value,
                pygame.Rect(box.x + 12, row_y, box.width - 24, 20),
                small=True,
                value_color=color,
            )
            row_y += 24

    def draw_cash_memo(self, panel: pygame.Rect, y: int, *, sold: bool = False) -> None:
        prop = self.current_property
        if not prop:
            return
        self.draw_text(UI_TEXT["cash_memo"], (panel.x + 18, y), 15, True, COLORS["muted"])
        y += 26
        lines = [
            UI_TEXT["cash_purchase_note"] if prop.status != "candidate" else UI_TEXT["after_strategy"],
            UI_TEXT["cash_after_sale_note"] if sold else UI_TEXT["cash_sale_note"],
        ]
        if sold:
            lines.append(f"{UI_TEXT['confirmed_profit']}: {yen(prop.sold_profit)}")
        for line in lines:
            self.draw_wrapped_text(
                f"・{line}",
                pygame.Rect(panel.x + 18, y, 316, 24),
                self.get_font(12),
                COLORS["muted"],
            )
            y += 23
        if not sold and prop.status in {"owned", "renovating", "listed"}:
            y += 2
            self.draw_key_value(UI_TEXT["cash_on_hand"], yen(self.cash), pygame.Rect(panel.x + 18, y, 316, 20), small=True)
            y += 22
            self.draw_key_value(UI_TEXT["holding_value"], yen(self.holding_value(prop)), pygame.Rect(panel.x + 18, y, 316, 20), small=True)
            y += 22
            self.draw_key_value(UI_TEXT["total_assets"], yen(self.total_assets_estimate()), pygame.Rect(panel.x + 18, y, 316, 20), small=True)
            y += 22
            self.draw_key_value(UI_TEXT["unrealized_profit"], self.signed_yen(self.unrealized_profit(prop)), pygame.Rect(panel.x + 18, y, 316, 20), small=True)
            if self.cash < LOW_CASH_WARNING:
                self.draw_wrapped_text(
                    UI_TEXT["cash_warning"],
                    pygame.Rect(panel.x + 18, y + 28, 316, 36),
                    self.get_font(12),
                    COLORS["red"],
                )

    def draw_sale_feedback(self, panel: pygame.Rect, y: int) -> None:
        prop = self.current_property
        if not prop:
            return
        reasons = prop.last_sale_feedback or self.sale_feedback_reasons(prop)
        self.draw_text(UI_TEXT["sale_feedback"], (panel.x + 18, y), 15, True, COLORS["muted"])
        y += 24
        for reason in reasons[:3]:
            self.draw_wrapped_text(
                f"・{reason}",
                pygame.Rect(panel.x + 18, y, 316, 30),
                self.get_font(11),
                COLORS["muted"],
            )
            y += 30

    def signed_yen(self, value: int) -> str:
        if value >= 0:
            return f"+{yen(value)}"
        return f"-{yen(abs(value))}"

    def signed_plain(self, value: int) -> str:
        if value > 0:
            return f"+{value}"
        return str(value)

    def draw_bottom_buttons(self) -> None:
        panel = pygame.Rect(24, 724, 1132, 72)
        self.draw_panel(panel, COLORS["panel"])
        labels = self.available_actions()
        if not labels:
            self.draw_wrapped_text(
                UI_TEXT["click_start"],
                panel.inflate(-40, -26),
                self.get_font(15),
                COLORS["muted"],
                center=True,
            )
        else:
            count = len(labels)
            gap = 12
            width = min(190, (panel.width - 40 - gap * (count - 1)) // max(1, count))
            total_width = width * count + gap * (count - 1)
            x = panel.centerx - total_width // 2
            for label, action, payload, enabled in labels:
                self.add_button(pygame.Rect(x, panel.y + 12, width, 48), label, action, payload, enabled)
                x += width + gap
        self.draw_buttons()

    def available_actions(self) -> list[tuple[str, str, Any, bool]]:
        prop = self.current_property
        if self.pending_confirmation == "sold":
            return [(UI_TEXT["confirm_next"], "confirm_next", None, True)]
        if self.pending_confirmation == "skipped":
            return [(UI_TEXT["confirm_next"], "confirm_next", None, True)]
        if self.pending_confirmation == "risk":
            return [(UI_TEXT["confirm_continue"], "confirm_continue", None, True)]
        if self.pending_confirmation == "appraisal_accepted":
            return [(UI_TEXT["confirm_continue"], "confirm_continue", None, True)]
        if self.pending_confirmation == "appraisal_rejected":
            return [(UI_TEXT["confirm_next"], "confirm_next", None, True)]
        if self.pending_confirmation == "purchase_offer":
            return [
                (UI_TEXT["accept_offer"], "accept_offer", None, True),
                (UI_TEXT["reject_offer"], "reject_offer", None, True),
            ]
        if not prop:
            return [(UI_TEXT["new_case"], "new_case", None, True)]
        if prop.status == "candidate":
            if self.is_purchase_banned(prop):
                return [(UI_TEXT["purchase_banned_skip"], "purchase_banned_skip", None, True)]
            if prop.appraisal_mode:
                return [
                    (f"{label}\n{yen(price)}", "appraisal_offer", price, True)
                    for label, price in self.appraisal_options(prop)
                ]
            return [
                (UI_TEXT["investigate"], "investigate", None, not prop.known_flags.get("investigated")),
                (UI_TEXT["hearing"], "hearing", None, not prop.known_flags.get("hearing")),
                (UI_TEXT["appraisal_offer"], "appraisal_start", None, True),
                (UI_TEXT["skip"], "skip", None, True),
            ]
        if prop.status == "skipped":
            return [(UI_TEXT["confirm_next"], "confirm_next", None, True)]
        if prop.status == "sold":
            return [(UI_TEXT["confirm_next"], "confirm_next", None, True)]
        if prop.status == "renovating":
            actions = [(UI_TEXT["next_day"], "next_day", None, True)]
            if prop.renovation_label_key and not prop.listed_price:
                for opt in SALE_OPTIONS:
                    price = self.preview_sale_price(prop, opt)
                    profit = self.projected_profit_for(prop, price) or 0
                    label = f"{UI_TEXT[str(opt['label_key'])]}\n{UI_TEXT['projected_profit']}: {yen(profit)}"
                    actions.append((label, "list_price", opt, True))
            elif prop.listed_price:
                actions.extend(
                    [
                        (UI_TEXT["price_down_500"], "discount", 500_000, True),
                        (UI_TEXT["price_down_1000"], "discount", 1_000_000, True),
                    ]
                )
                if self.can_loss_cut(prop):
                    actions.append((UI_TEXT["loss_cut"], "loss_cut", None, True))
            return actions
        if self.is_sale_active(prop):
            actions = [
                (UI_TEXT["next_day"], "next_day", None, True),
                (UI_TEXT["price_down_500"], "discount", 500_000, True),
                (UI_TEXT["price_down_1000"], "discount", 1_000_000, True),
            ]
            if self.can_loss_cut(prop):
                actions.append((UI_TEXT["loss_cut"], "loss_cut", None, True))
            return actions
        if prop.status == "owned" and not prop.renovation_label_key:
            return [(UI_TEXT[str(opt["label_key"])], "renovation", opt, True) for opt in RENOVATION_OPTIONS]
        if prop.status == "owned" and prop.renovation_label_key:
            actions = []
            for opt in SALE_OPTIONS:
                price = self.preview_sale_price(prop, opt)
                profit = self.projected_profit_for(prop, price) or 0
                label = f"{UI_TEXT[str(opt['label_key'])]}\n{UI_TEXT['projected_profit']}: {yen(profit)}"
                actions.append((label, "list_price", opt, True))
            return actions
        return []

    def visible_risk_rows(self, prop: PropertyCase) -> list[tuple[str, str, bool]]:
        rows: list[tuple[str, str, bool]] = []
        checks = (
            ("seller_antisocial", self.seller_antisocial_text(prop)),
            ("rebuildable", UI_TEXT["possible"] if prop.rebuildable else UI_TEXT["not_possible"]),
            ("road_access_good", UI_TEXT["good"] if prop.road_access_good else UI_TEXT["bad"]),
            ("illegal_building", self.yes_no(prop.illegal_building)),
            ("rain_leak_level", self.risk_level_text(prop.rain_leak_level)),
            ("building_tilt_level", self.risk_level_text(prop.building_tilt_level)),
            ("termite_damage", self.termite_level_text(prop.termite_level)),
            ("leftover_items", self.leftover_level_text(prop.leftover_level)),
            ("nearby_antisocial_base", self.yes_no(prop.nearby_antisocial_base)),
            ("neighbor_garbage_house", self.yes_no(prop.neighbor_garbage_house)),
            ("flood_damage", self.yes_no(prop.flood_damage)),
            ("incident_property", self.yes_no(prop.incident_property)),
            ("neighbor_trouble", self.yes_no(prop.neighbor_trouble)),
        )
        for key, value in checks:
            rows.append((RISK_LABELS[key], value, prop.known_flags.get(key, False)))
            if key == "rebuildable":
                reason = prop.rebuild_blocker_reason if prop.rebuild_blocker_reason else UI_TEXT["none"]
                rows.append((UI_TEXT["reason"], reason, prop.known_flags.get(key, False)))
        return rows

    def preview_sale_price(self, prop: PropertyCase, option: dict[str, Any]) -> int:
        target = 2_500_000 if prop.renovation_budget > 0 else 1_500_000
        price = self.calculate_total_cost(prop) + target + int(option["profit_delta"])
        return max(1_000_000, round(price / 100_000) * 100_000)

    def station_distance_text(self, prop: PropertyCase) -> str:
        return f"徒歩{prop.station_walk_minutes}分 / {prop.station_distance_m}m"

    def yes_no(self, value: bool) -> str:
        return UI_TEXT["yes"] if value else UI_TEXT["no"]

    def risk_level_text(self, value: str) -> str:
        if value == "suspected":
            return UI_TEXT["risk_suspected"]
        if value == "confirmed":
            return UI_TEXT["risk_confirmed"]
        return UI_TEXT["risk_none"]

    def termite_level_text(self, value: str) -> str:
        return self.risk_level_text(value)

    def leftover_level_text(self, value: str) -> str:
        if value == "small":
            return UI_TEXT["risk_small"]
        if value == "many":
            return UI_TEXT["risk_many"]
        return UI_TEXT["risk_none"]

    def seller_antisocial_text(self, prop: PropertyCase) -> str:
        return self.risk_level_text(prop.seller_antisocial_level)

    def draw_panel(self, rect: pygame.Rect, color: tuple[int, int, int]) -> None:
        shadow = rect.move(0, 2)
        pygame.draw.rect(self.screen, COLORS["shadow"], shadow, border_radius=8)
        pygame.draw.rect(self.screen, color, rect, border_radius=8)
        pygame.draw.rect(self.screen, COLORS["line"], rect, 1, border_radius=8)

    def draw_stat_chip(self, label: str, value: str, rect: pygame.Rect) -> None:
        pygame.draw.rect(self.screen, COLORS["soft"], rect, border_radius=8)
        self.draw_text(label, (rect.x + 10, rect.y + 8), 12, False, COLORS["muted"])
        value_surface = self.get_font(14, True).render(value, True, COLORS["text"])
        self.screen.blit(value_surface, value_surface.get_rect(midright=(rect.right - 10, rect.y + 18)))

    def draw_status_badge(self, prop: PropertyCase, rect: pygame.Rect) -> None:
        color = COLORS["soft_green"]
        if prop.status == "sold":
            color = COLORS["yellow"]
        elif prop.status == "skipped":
            color = COLORS["button_disabled"]
        elif prop.status == "listed":
            color = (220, 234, 246)
        elif prop.status == "renovating":
            color = (249, 230, 207)
        pygame.draw.rect(self.screen, color, rect, border_radius=8)
        pygame.draw.rect(self.screen, COLORS["line"], rect, 1, border_radius=8)
        label_key = "renovating_listed" if prop.status == "renovating" and prop.listed_price else prop.status
        text = self.get_font(13, True).render(UI_TEXT[label_key], True, COLORS["text"])
        self.screen.blit(text, text.get_rect(center=rect.center))

    def draw_key_value(
        self,
        label: str,
        value: str,
        rect: pygame.Rect,
        *,
        small: bool = False,
        value_color: tuple[int, int, int] | None = None,
    ) -> None:
        font_size = 13 if small else 14
        label_font = self.get_font(font_size)
        value_font = self.get_font(font_size, True)
        label_surf = label_font.render(label, True, COLORS["muted"])
        self.screen.blit(label_surf, (rect.x, rect.y + 3))
        value_surf = value_font.render(value, True, value_color or COLORS["text"])
        value_rect = value_surf.get_rect(midright=(rect.right, rect.y + rect.height // 2 + 1))
        self.screen.blit(value_surf, value_rect)
        pygame.draw.line(
            self.screen,
            COLORS["line"],
            (rect.x, rect.bottom + 2),
            (rect.right, rect.bottom + 2),
            1,
        )

    def draw_text(
        self,
        text: str,
        pos: tuple[int, int],
        size: int,
        bold: bool = False,
        color: tuple[int, int, int] | None = None,
    ) -> None:
        surf = self.get_font(size, bold).render(text, True, color or COLORS["text"])
        self.screen.blit(surf, pos)

    def draw_wrapped_text(
        self,
        text: str,
        rect: pygame.Rect,
        font: pygame.font.Font,
        color: tuple[int, int, int],
        *,
        center: bool = False,
        line_gap: int = 4,
    ) -> None:
        lines = self.wrap_text(text, font, rect.width)
        line_height = font.get_linesize() + line_gap
        total_h = len(lines) * line_height
        y = rect.y + max(0, (rect.height - total_h) // 2) if center else rect.y
        for line in lines:
            surf = font.render(line, True, color)
            x = rect.x + (rect.width - surf.get_width()) // 2 if center else rect.x
            self.screen.blit(surf, (x, y))
            y += line_height
            if y > rect.bottom:
                break

    def wrap_text(self, text: str, font: pygame.font.Font, max_width: int) -> list[str]:
        if not text:
            return [""]
        lines: list[str] = []
        current = ""
        for char in text:
            if char == "\n":
                lines.append(current)
                current = ""
                continue
            candidate = current + char
            if font.size(candidate)[0] <= max_width or not current:
                current = candidate
            else:
                lines.append(current)
                current = char
        if current:
            lines.append(current)
        return lines

    def add_button(
        self,
        rect: pygame.Rect,
        text: str,
        action: str,
        payload: Any = None,
        enabled: bool = True,
    ) -> None:
        self.buttons.append(Button(rect, text, action, payload, enabled))

    def draw_buttons(self) -> None:
        for button in self.buttons:
            hovered = button.rect.collidepoint(self.mouse_pos) and button.enabled
            bg = COLORS["button_hover"] if hovered else COLORS["button"]
            if not button.enabled:
                bg = COLORS["button_disabled"]
            border = COLORS["green"] if hovered else COLORS["line_dark"]
            pygame.draw.rect(self.screen, bg, button.rect, border_radius=8)
            pygame.draw.rect(self.screen, border, button.rect, 1, border_radius=8)
            color = COLORS["green_dark"] if button.enabled else COLORS["muted"]
            font = self.get_font(15, True)
            lines = self.wrap_text(button.text, font, button.rect.width - 18)
            line_height = font.get_linesize()
            y = button.rect.centery - (len(lines) * line_height) // 2
            for line in lines[:2]:
                surf = font.render(line, True, color)
                self.screen.blit(surf, surf.get_rect(center=(button.rect.centerx, y + line_height // 2)))
                y += line_height

    def draw_shimarisu_image(
        self,
        image: pygame.Surface | None,
        rect: pygame.Rect,
        *,
        large: bool = False,
        animated: bool = False,
    ) -> None:
        offset_x, offset_y = self.shimarisu_motion_offset() if animated else (0, 0)
        center = (rect.centerx + offset_x, rect.centery + offset_y)
        if image:
            fitted = fit_surface(image, (rect.width - 12, rect.height - 12))
            image_rect = fitted.get_rect(center=center)
            self.screen.blit(fitted, image_rect)
            return

        pygame.draw.rect(self.screen, COLORS["soft"], rect, border_radius=8)
        pygame.draw.rect(self.screen, COLORS["line"], rect, 1, border_radius=8)
        body_color = COLORS["orange"]
        cx, cy = center
        pygame.draw.circle(self.screen, body_color, (cx, cy - 10), rect.width // (4 if large else 5))
        pygame.draw.circle(self.screen, (255, 232, 169), (cx, cy + 8), rect.width // (5 if large else 6))
        pygame.draw.circle(self.screen, COLORS["text"], (cx - 18, cy - 16), 5)
        pygame.draw.circle(self.screen, COLORS["text"], (cx + 18, cy - 16), 5)
        pygame.draw.arc(
            self.screen,
            body_color,
            pygame.Rect(cx + 28, cy - 22, 54, 64),
            math.radians(250),
            math.radians(95),
            8,
        )

    def shimarisu_motion_offset(self) -> tuple[int, int]:
        prop = self.current_property
        ticks = pygame.time.get_ticks()
        bob = int(math.sin(ticks / 360) * 3)
        side = 0
        if prop:
            if prop.status == "listed":
                side = int(math.sin(ticks / 230) * 4)
            elif prop.status == "sold":
                side = 5
                bob = int(math.sin(ticks / 420) * 2) - 1
            elif prop.status == "renovating":
                side = int(math.sin(ticks / 300) * 2)
        return side, bob

    def draw_office_area(self, rect: pygame.Rect) -> None:
        pygame.draw.rect(self.screen, COLORS["panel_alt"], rect, border_radius=8)
        pygame.draw.rect(self.screen, COLORS["line"], rect, 1, border_radius=8)
        self.draw_text(UI_TEXT["office"], (rect.x + 10, rect.y + 8), 13, True, COLORS["muted"])

        floor_y = rect.bottom - 20
        pygame.draw.line(self.screen, COLORS["line_dark"], (rect.x + 12, floor_y), (rect.right - 12, floor_y), 2)

        desk = pygame.Rect(rect.x + 26, rect.y + 76, 84, 34)
        pygame.draw.rect(self.screen, (170, 119, 78), desk, border_radius=4)
        pygame.draw.rect(self.screen, (126, 83, 54), desk, 1, border_radius=4)
        pygame.draw.line(self.screen, (126, 83, 54), (desk.x + 14, desk.bottom), (desk.x + 14, floor_y), 3)
        pygame.draw.line(self.screen, (126, 83, 54), (desk.right - 14, desk.bottom), (desk.right - 14, floor_y), 3)
        pygame.draw.rect(self.screen, COLORS["white"], pygame.Rect(desk.x + 12, desk.y - 16, 26, 16), border_radius=2)
        pygame.draw.rect(self.screen, COLORS["line_dark"], pygame.Rect(desk.x + 12, desk.y - 16, 26, 16), 1, border_radius=2)

        if self.sold_count >= 1:
            pot = pygame.Rect(rect.x + 132, rect.y + 91, 22, 18)
            pygame.draw.rect(self.screen, (169, 111, 82), pot, border_radius=4)
            pygame.draw.circle(self.screen, COLORS["green"], (pot.centerx - 6, pot.y - 8), 9)
            pygame.draw.circle(self.screen, COLORS["green_dark"], (pot.centerx + 7, pot.y - 12), 8)

        if self.sold_count >= 2:
            shelf = pygame.Rect(rect.x + 170, rect.y + 55, 42, 60)
            pygame.draw.rect(self.screen, (190, 150, 96), shelf, border_radius=4)
            pygame.draw.rect(self.screen, (126, 83, 54), shelf, 1, border_radius=4)
            for line_y in (shelf.y + 18, shelf.y + 36):
                pygame.draw.line(self.screen, (126, 83, 54), (shelf.x + 4, line_y), (shelf.right - 4, line_y), 1)
            pygame.draw.rect(self.screen, COLORS["yellow"], pygame.Rect(shelf.x + 7, shelf.y + 6, 8, 10))
            pygame.draw.rect(self.screen, COLORS["soft_green"], pygame.Rect(shelf.x + 20, shelf.y + 24, 9, 10))

        if self.cash >= 25_000_000:
            sign = pygame.Rect(rect.x + 24, rect.y + 34, 90, 24)
            pygame.draw.rect(self.screen, COLORS["soft_green"], sign, border_radius=4)
            pygame.draw.rect(self.screen, COLORS["green"], sign, 1, border_radius=4)
            text = self.get_font(11, True).render("DAKE", True, COLORS["green_dark"])
            self.screen.blit(text, text.get_rect(center=sign.center))

        if self.reputation >= 60:
            table = pygame.Rect(rect.x + 126, rect.y + 72, 36, 22)
            pygame.draw.ellipse(self.screen, (207, 175, 122), table)
            pygame.draw.ellipse(self.screen, (126, 83, 54), table, 1)
            pygame.draw.circle(self.screen, COLORS["soft_green"], (table.x - 10, table.centery), 8)
            pygame.draw.circle(self.screen, COLORS["soft_green"], (table.right + 10, table.centery), 8)

        prop = self.current_property
        if prop and prop.status == "listed":
            sale_sign = pygame.Rect(rect.x + 132, rect.y + 34, 70, 24)
            pygame.draw.rect(self.screen, COLORS["yellow"], sale_sign, border_radius=4)
            pygame.draw.rect(self.screen, COLORS["orange"], sale_sign, 1, border_radius=4)
            text = self.get_font(10, True).render("SALE", True, COLORS["text"])
            self.screen.blit(text, text.get_rect(center=sale_sign.center))
        if prop and prop.status == "sold":
            sparkle_x = rect.right - 34
            sparkle_y = rect.y + 35 + int(math.sin(pygame.time.get_ticks() / 180) * 2)
            pygame.draw.line(self.screen, COLORS["orange"], (sparkle_x, sparkle_y - 9), (sparkle_x, sparkle_y + 9), 2)
            pygame.draw.line(self.screen, COLORS["orange"], (sparkle_x - 9, sparkle_y), (sparkle_x + 9, sparkle_y), 2)
            pygame.draw.circle(self.screen, COLORS["yellow"], (sparkle_x, sparkle_y), 4)


def main() -> None:
    ShimarisuRealEstateGame().run()


if __name__ == "__main__":
    main()
