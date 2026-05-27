def print_gugudan(n: int, start: int = 1, end: int = 9) -> None:
    print(f"\n=== {n}단 ===")
    for i in range(start, end + 1):
        print(f"{n} x {i} = {n * i}")


def print_all_gugudan(start: int = 2, end: int = 9) -> None:
    for n in range(start, end + 1):
        print_gugudan(n)


def run():
    print("구구단 프로그램")
    print("1) 특정 단 출력")
    print("2) 전체 2단~9단 출력")
    choice = input("선택하세요 (1/2): ").strip()

    if choice == "1":
        try:
            n = int(input("출력할 단을 입력하세요 (2~9): ").strip())
        except ValueError:
            print("숫자를 입력해주세요.")
            return
        if n < 2 or n > 9:
            print("2에서 9 사이의 숫자를 입력해주세요.")
            return
        print_gugudan(n)
    elif choice == "2":
        print_all_gugudan()
    else:
        print("올바른 선택이 아닙니다.")


if __name__ == "__main__":
    run()
