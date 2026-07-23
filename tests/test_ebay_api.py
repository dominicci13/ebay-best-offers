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
    assert script._token_env("Account One") == "EBAY_AUTH_TOKEN_ACCOUNTONE"
    assert script._token_env("account-two") == "EBAY_AUTH_TOKEN_ACCOUNTTWO"
    assert script._token_env("Acct-3 Intl") == "EBAY_AUTH_TOKEN_ACCT3INTL"       # spaces, hyphen, digit all stripped


def test_build_xml_has_token_and_active_status():
    xml = script.build_get_best_offers_xml("ABC123")
    assert "GetBestOffersRequest" in xml
    assert "<eBayAuthToken>ABC123</eBayAuthToken>" in xml
    assert "<BestOfferStatus>Active</BestOfferStatus>" in xml


def test_build_xml_paginates_and_defaults_to_page_one():
    assert "<PageNumber>1</PageNumber>" in script.build_get_best_offers_xml("T")
    assert "<PageNumber>3</PageNumber>" in script.build_get_best_offers_xml("T", 3)


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
    assert result["total_pages"] == 1                  # no PaginationResult -> single page


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


def _page_xml(item_id: str, page: int, total_pages: int) -> str:
    """A one-item GetBestOffers page carrying pagination metadata."""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<GetBestOffersResponse xmlns="urn:ebay:apis:eBLBaseComponents">'
        "<Ack>Success</Ack>"
        f"<PaginationResult><TotalNumberOfPages>{total_pages}</TotalNumberOfPages>"
        "<TotalNumberOfEntries>2</TotalNumberOfEntries></PaginationResult>"
        f"<PageNumber>{page}</PageNumber>"
        "<ItemBestOffersArray><ItemBestOffers>"
        f"<Item><ItemID>{item_id}</ItemID></Item>"
        "<BestOffersArray><BestOffer>"
        f"<BestOfferID>bo-{item_id}</BestOfferID>"
        '<Price currencyID="USD">100.0</Price>'
        "<BestOfferCodeType>BuyerBestOffer</BestOfferCodeType><Status>Pending</Status>"
        "</BestOffer></BestOffersArray></ItemBestOffers></ItemBestOffersArray>"
        "</GetBestOffersResponse>"
    )


def test_parse_reads_total_pages():
    assert script.parse_best_offers(_page_xml("111", 1, 2))["total_pages"] == 2


class _FakeResp:
    def __init__(self, content: str):
        self.content = content.encode("utf-8")
        self.status_code = 200


def test_get_best_offers_walks_every_page(monkeypatch):
    """The account read must follow pagination — page 1 alone drops later offers."""
    pages = [_page_xml("111111111111", 1, 2), _page_xml("222222222222", 2, 2)]
    seen_pages = []

    def fake_post(url, data=None, headers=None, timeout=None):
        body = data.decode("utf-8")
        page = int(body.split("<PageNumber>")[1].split("</PageNumber>")[0])
        seen_pages.append(page)
        return _FakeResp(pages[page - 1])

    monkeypatch.setattr(script, "get_env", lambda *a, **k: "x")   # header creds
    monkeypatch.setattr(script.requests, "post", fake_post)

    offers = script.get_best_offers("TKN")
    assert seen_pages == [1, 2]                                   # both pages requested, in order
    assert {o["item_number"] for o in offers} == {"111111111111", "222222222222"}


def test_get_best_offers_single_page_stops(monkeypatch):
    def fake_post(url, data=None, headers=None, timeout=None):
        return _FakeResp(_page_xml("111111111111", 1, 1))       # total_pages = 1

    monkeypatch.setattr(script, "get_env", lambda *a, **k: "x")
    monkeypatch.setattr(script.requests, "post", fake_post)
    offers = script.get_best_offers("TKN")
    assert len(offers) == 1                                       # no extra page fetched
