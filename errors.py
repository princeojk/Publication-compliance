
from dataclasses import dataclass

@dataclass
class ComplianceChecker:
    rule: str
    expected: str | bool | float | dict[str, float]
    actual: str | bool | float | None | dict[str, float] = None

    def __str__(self) -> str:
        return f"{self.rule}: expected {self.expected}, found: {self.actual}"

    def serilizer(self) -> dict[str, str|bool|float|None|dict[str, float]]:
        return {
            'rule': self.rule,
            'expected': self.expected,
            'actual': self.actual
        }

