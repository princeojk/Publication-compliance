from dataclasses import dataclass

@dataclass
class SpaceAttributes:
    before: float
    after: float

    def __str__(self) -> str:
        return f"before: {self.before}: after {self.after}"

    def serializer(self) -> dict[str, float]:

        return {
            'before': self.before,
            'after': self.after
        }