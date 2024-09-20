import flet as ft



def main(page: ft.Page):
    page.add(ft.SafeArea(ft.Text("Hello, Flet!")))
    page.add(MyButton(text="OK"), MyButton(text="Cancel"))


ft.app(main)