class SpreadsheetRow:
    original_name: str
    description: str
    skip: bool

    def __init__(self, original_name: str, description: str, skip: str) -> None:
        self.original_name = original_name
        self.description = description
        self.skip = self.should_skip(skip)

    def __repr__(self) -> str:
        return f"SpreadsheetRow({self.original_name}, {self.description}, {self.skip})"

    def should_skip(self, skip: str) -> bool:
        if skip:
            if skip.lower() == "y":
                return True
        return False

