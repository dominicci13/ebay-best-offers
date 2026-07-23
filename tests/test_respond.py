"""Tests for the eBay Trading API acting layer (RespondToBestOffer + orchestrator).

All offline — no network. The orchestrator's live path is exercised by
monkeypatching :func:`respond_to_best_offer`, so no call ever reaches eBay.
"""
from __future__ import annotations

import pandas as pd

import run_ebay_best_offers as script

SETTINGS = {"commission": 0.12, "min_discount": 0.05, "max_discount": 0.10}

# One row per offer, incl. a skip reason (Expired) and a missing-id counter. Each
# row carries its best_offer_id + offer_quantity (what respond_to_offers reads) plus
# the numeric columns that feed the per-offer log line (_offer_log_line).
RESULTS = pd.DataFrame(
    [
        # account, item, action, cx_offer, current_price, counter, counter_margin, buyer_margin, total_cost, best_offer_id, offer_quantity
        ("Acct1", "111", "Accepted",      95.00, 100.00,   None,  None, 0.10, 80.00,  "b-accept",  1),
        ("Acct1", "222", "Counteroffer", 150.00, 200.00, 192.39, 0.09,  None, 120.00, "b-counter", 3),
        ("Acct1", "333", "Declined",      50.00, 100.00,   None,  None, None, 90.00,  "b-decline", 1),
        ("Acct1", "444", "Expired Offer",  0.00,  50.00,   None,  None, None, 40.00,  "b-expired", 1),
        ("Acct1", "555", "Counteroffer",  80.00, 120.00,  88.00, 0.09,  None, 70.00,  None,        1),
    ],
    columns=["account", "item_number", "action", "cx_offer", "current_price",
             "counter", "counter_margin", "buyer_margin", "total_cost",
             "best_offer_id", "offer_quantity"],
)


def _row(item_number):
    """The RESULTS row for an item as a namedtuple (what respond_to_offers iterates)."""
    return next(RESULTS[RESULTS.item_number == item_number].itertuples(index=False))


# --- build_respond_offer_xml -------------------------------------------------

def test_accept_xml_omits_price_and_message():
    xml = script.build_respond_offer_xml("TKN", "111", "b1", "Accept")
    assert "<Action>Accept</Action>" in xml
    assert "<ItemID>111</ItemID>" in xml and "<BestOfferID>b1</BestOfferID>" in xml
    assert "CounterOfferPrice" not in xml
    assert "SellerResponse" not in xml


def test_counter_xml_has_price_quantity_and_message():
    xml = script.build_respond_offer_xml(
        "TKN", "222", "b2", "Counter", counter_price=192.39, counter_quantity=3, message="hi"
    )
    assert "<Action>Counter</Action>" in xml
    assert '<CounterOfferPrice currencyID="USD">192.39</CounterOfferPrice>' in xml
    assert "<CounterOfferQuantity>3</CounterOfferQuantity>" in xml
    assert "<SellerResponse>hi</SellerResponse>" in xml


def test_decline_xml_carries_message_but_no_price():
    xml = script.build_respond_offer_xml("TKN", "333", "b3", "Decline", message="sorry")
    assert "<Action>Decline</Action>" in xml
    assert "CounterOfferPrice" not in xml
    assert "<SellerResponse>sorry</SellerResponse>" in xml


def test_token_and_message_are_escaped():
    xml = script.build_respond_offer_xml("a&b", "1", "b", "Decline", message="x<y&z")
    assert "<eBayAuthToken>a&amp;b</eBayAuthToken>" in xml
    assert "x&lt;y&amp;z" in xml


def test_message_hard_capped_at_limit():
    long_msg = "z" * 400
    xml = script.build_respond_offer_xml("T", "1", "b", "Decline", message=long_msg)
    body = xml.split("<SellerResponse>")[1].split("</SellerResponse>")[0]
    assert len(body) == script.MESSAGE_LIMIT


def test_canned_messages_fit_the_limit():
    assert len(script.DECLINE_MESSAGE) <= script.MESSAGE_LIMIT
    # Counter with a wide, comma-grouped price still fits.
    assert len(script.COUNTER_MESSAGE.format(price="99,999.99")) <= script.MESSAGE_LIMIT


# --- parse_respond_response --------------------------------------------------

def test_parse_success():
    xml = """<?xml version="1.0" encoding="UTF-8"?>
    <RespondToBestOfferResponse xmlns="urn:ebay:apis:eBLBaseComponents">
      <Ack>Success</Ack>
    </RespondToBestOfferResponse>"""
    result = script.parse_respond_response(xml)
    assert result["ack"] == "Success"
    assert result["errors"] == []


def test_parse_failure_surfaces_errors():
    xml = """<?xml version="1.0" encoding="UTF-8"?>
    <RespondToBestOfferResponse xmlns="urn:ebay:apis:eBLBaseComponents">
      <Ack>Failure</Ack>
      <Errors><ErrorCode>21916</ErrorCode><LongMessage>Best Offer no longer available.</LongMessage></Errors>
    </RespondToBestOfferResponse>"""
    result = script.parse_respond_response(xml)
    assert result["ack"] == "Failure"
    assert any("21916" in e or "no longer available" in e for e in result["errors"])


# --- respond_to_offers (orchestrator) ----------------------------------------

def test_dry_run_sends_nothing(monkeypatch):
    def boom(*a, **k):
        raise AssertionError("respond_to_best_offer must not be called in a dry run")
    monkeypatch.setattr(script, "respond_to_best_offer", boom)
    acted = script.respond_to_offers(RESULTS, "TKN", SETTINGS, live=False)
    assert acted == {}                                 # nothing answered, nothing stamped


def test_live_maps_actions_and_skips_non_actionable(monkeypatch):
    calls = []

    def fake(token, item_id, best_offer_id, action, counter_price=None, counter_quantity=1, message=None):
        calls.append({"item": item_id, "action": action, "price": counter_price,
                      "qty": counter_quantity, "msg": message, "boid": best_offer_id})
        return {"ack": "Success", "errors": []}

    monkeypatch.setattr(script, "respond_to_best_offer", fake)
    acted = script.respond_to_offers(RESULTS, "TKN", SETTINGS, live=True)

    # Accept, Counter, Decline answered; Expired skipped; missing-id counter skipped.
    assert set(acted) == {"b-accept", "b-counter", "b-decline"}   # stamped by BestOfferID
    sent = {c["item"]: c for c in calls}
    assert set(sent) == {"111", "222", "333"}
    assert sent["111"]["action"] == "Accept" and sent["111"]["price"] is None and sent["111"]["msg"] is None
    assert sent["222"]["action"] == "Counter" and sent["222"]["price"] == 192.39
    assert sent["222"]["qty"] == 3                      # quantity carried from the API read
    assert "192.39" in sent["222"]["msg"]              # counter price filled into the message
    assert sent["333"]["action"] == "Decline" and sent["333"]["price"] is None
    assert sent["333"]["msg"] == script.DECLINE_MESSAGE


def test_live_bad_ack_is_not_stamped(monkeypatch):
    def fake(*a, **k):
        return {"ack": "Failure", "errors": ["21916: Best Offer no longer available."]}
    monkeypatch.setattr(script, "respond_to_best_offer", fake)
    acted = script.respond_to_offers(RESULTS, "TKN", SETTINGS, live=True)
    assert acted == {}                                 # a rejected offer is never marked answered


def test_two_offers_on_one_item_are_both_answered(monkeypatch):
    """A listing with two buyer offers answers both, not just one (the multi-offer fix)."""
    calls = []

    def fake(token, item_id, best_offer_id, action, counter_price=None, counter_quantity=1, message=None):
        calls.append(best_offer_id)
        return {"ack": "Success", "errors": []}

    monkeypatch.setattr(script, "respond_to_best_offer", fake)
    two = pd.DataFrame(
        [
            ("Acct1", "900", "Counteroffer", 50.0, 100.0, 92.0, 0.09, None, 70.0, "boid-A", 1),
            ("Acct1", "900", "Declined",     30.0, 100.0, None, None, None, 90.0, "boid-B", 1),
        ],
        columns=RESULTS.columns,
    )
    acted = script.respond_to_offers(two, "TKN", SETTINGS, live=True)
    assert set(acted) == {"boid-A", "boid-B"}          # both offers on item 900 answered
    assert calls == ["boid-A", "boid-B"]


# --- _offer_log_line (the per-offer log wording) -----------------------------

def test_log_line_accept():
    assert script._offer_log_line(_row("111"), SETTINGS) == (
        "Offer accepted for item 111 at $95.00 with a profit of 10.00%."
    )


def test_log_line_counter():
    assert script._offer_log_line(_row("222"), SETTINGS) == (
        "Counteroffered for item 222 from $150.00 to $192.39 with a profit of 9.00%."
    )


def test_log_line_decline_shows_best_case_margin_and_lowest_discount():
    # 100 * (1 - 0.05) = 95; margin(95, 90, 0.12) = (95 - 90 - 11.4)/95 = -6.74% -> negative.
    assert script._offer_log_line(_row("333"), SETTINGS) == (
        "Offer declined for item 333. Margin was negative at -6.74% "
        "after the lowest discount 5% applied."
    )
