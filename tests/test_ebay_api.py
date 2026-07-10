"""Tests for the eBay Trading API read layer (GetBestOffers build + parse)."""
from __future__ import annotations

import run_ebay_best_offers as script

# A realistic GetBestOffers response: namespaced, two items, a buyer UserID that
# our parser must NOT surface (keeps the account-deletion exemption valid).
SUCCESS_XML = """<?xml version="1.0" encoding="UTF-8"?>
<GetBestOffersResponse xmlns="urn:ebay:apis:eBLBaseComponents">
  <Ack>Success</Ack>
  <ItemBestOffersArray>
    <ItemBestOffers>
      <Item><ItemID>111111111111</ItemID></Item>
      <BestOffersArray>
        <BestOffer>
          <BestOfferID>900001</BestOfferID>
          <Buyer><UserID>somebuyer</UserID></Buyer>
          <Price currencyID="USD">130.0</Price>
          <Quantity>2</Quantity>
          <BestOfferCodeType>BuyerBestOffer</BestOfferCodeType>
          <Status>Pending</Status>
        </BestOffer>
      </BestOffersArray>
    </ItemBestOffers>
    <ItemBestOffers>
      <Item><ItemID>222222222222</ItemID></Item>
      <BestOffersArray>
        <BestOffer>
          <BestOfferID>900002</BestOfferID>
          <Price currencyID="USD">45.50</Price>
          <BestOfferCodeType>BuyerBestOffer</BestOfferCodeType>
          <Status>Pending</Status>
        </BestOffer>
      </BestOffersArray>
    </ItemBestOffers>
  </ItemBestOffersArray>
</GetBestOffersResponse>"""

FAILURE_XML = """<?xml version="1.0" encoding="UTF-8"?>
<GetBestOffersResponse xmlns="urn:ebay:apis:eBLBaseComponents">
  <Ack>Failure</Ack>
  <Errors>
    <ErrorCode>931</ErrorCode>
    <SeverityCode>Error</SeverityCode>
    <LongMessage>Auth token is invalid.</LongMessage>
  </Errors>
</GetBestOffersResponse>"""


def test_token_env_derives_names():
    assert script._token_env("Account1") == "EBAY_AUTH_TOKEN_ACCOUNT1"
    assert script._token_env("Account2") == "EBAY_AUTH_TOKEN_ACCOUNT2"
    assert script._token_env("Account4") == "EBAY_AUTH_TOKEN_ACCOUNT4"


def test_build_xml_has_token_and_active_status():
    xml = script.build_get_best_offers_xml("ABC123")
    assert "GetBestOffersRequest" in xml
    assert "<eBayAuthToken>ABC123</eBayAuthToken>" in xml
    assert "<BestOfferStatus>Active</BestOfferStatus>" in xml


def test_build_xml_escapes_token():
    xml = script.build_get_best_offers_xml("a&b<c")
    assert "a&amp;b&lt;c" in xml
    assert "<eBayAuthToken>a&b<c" not in xml


def test_parse_success_extracts_offers():
    result = script.parse_best_offers(SUCCESS_XML)
    assert result["ack"] == "Success"
    assert result["errors"] == []
    by_item = {o["item_number"]: o for o in result["offers"]}
    assert len(by_item) == 2
    a = by_item["111111111111"]
    assert a["cx_offer"] == 130.0
    assert a["best_offer_id"] == "900001"
    assert a["quantity"] == 2                          # read from <Quantity>
    assert a["code"] == "BuyerBestOffer"
    assert a["status"] == "Pending"
    assert by_item["222222222222"]["cx_offer"] == 45.5
    assert by_item["222222222222"]["quantity"] == 1    # no <Quantity> -> defaults to 1


def test_parse_never_surfaces_buyer_identity():
    result = script.parse_best_offers(SUCCESS_XML)
    for offer in result["offers"]:
        assert "somebuyer" not in str(offer)          # buyer UserID never captured
        assert set(offer) == {"item_number", "cx_offer", "best_offer_id", "quantity", "code", "status"}


def test_parse_failure_returns_errors_and_no_offers():
    result = script.parse_best_offers(FAILURE_XML)
    assert result["ack"] == "Failure"
    assert result["offers"] == []
    assert any("931" in e or "Auth token" in e for e in result["errors"])


def test_parse_accepts_bytes_with_encoding_declaration():
    result = script.parse_best_offers(SUCCESS_XML.encode("utf-8"))
    assert len(result["offers"]) == 2
