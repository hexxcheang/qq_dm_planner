"""Plain data structures shared by the parser, generator and UI."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class LineItem:
    description: str
    qty: float
    units: str = "Pcs"
    unit_price: float = 0.0

    @property
    def amount(self) -> float:
        return round(self.qty * self.unit_price, 2)


@dataclass
class Customer:
    name: str = ""
    address1: str = ""
    address2: str = ""
    attention: str = ""

    def lines(self) -> list[str]:
        return [self.name, self.address1, self.address2, self.attention]


@dataclass
class PurchaseOrder:
    po_number: str = ""
    items: list[LineItem] = field(default_factory=list)
    customer: Customer | None = None  # None -> keep the customer in the template

    @property
    def total(self) -> float:
        return round(sum(i.amount for i in self.items), 2)


def combine_purchase_orders(pos: list[PurchaseOrder]) -> PurchaseOrder:
    """One PO holding every item of ``pos`` (in order), for a single invoice and DO.

    The PO numbers are listed together ("POR13495, POR13557"); a number that
    appears twice is listed once. Items are kept as separate lines, so each
    stays traceable to its PO.
    """
    numbers: list[str] = []
    for po in pos:
        number = po.po_number.strip()
        if number and number not in numbers:
            numbers.append(number)
    items = [LineItem(i.description, i.qty, i.units, i.unit_price) for po in pos for i in po.items]
    return PurchaseOrder(po_number=", ".join(numbers), items=items)
