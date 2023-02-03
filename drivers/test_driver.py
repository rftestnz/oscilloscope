from base_scope_driver import ScopeDriver


class TestScope(ScopeDriver):
    def __init__(self) -> None:
        super().__init__()

    def write(self, command: str) -> None:
        return super().write(command)

    def read(self) -> str:
        return super().read()


if __name__ == "__main__":

    driver_test = TestScope()
