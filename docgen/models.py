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
