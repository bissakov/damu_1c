import dataclasses
import logging
import os
import re
import sys
import warnings
from contextlib import suppress
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from time import sleep
from typing import Optional, Generator, cast

import dotenv
import pyperclip
from pywinauto import ElementNotFoundError
from pywinauto.keyboard import send_keys
from urllib3.exceptions import InsecureRequestWarning

from src.utils.automation import (
    child_win,
    a,
    UiaWindow,
    wait_for,
    window,
    click_type_keys,
    contains_text,
    UiaButton,
    text,
    text_to_float,
    check,
    UiaPane,
    UiaList,
    children,
    UiaListItem,
    iter_children,
)

project_folder = Path(__file__).resolve().parent.parent
os.environ["project_folder"] = str(project_folder)
os.chdir(project_folder)
sys.path.append(str(project_folder))

from src.utils.app import App
from src.utils.db_manager import DatabaseManager

from src.utils import logger


def prepare_query(contragent: str) -> str:
    query = f"""
        ВЫБРАТЬ Проекты.Ссылка
        ИЗ Справочник.Контрагенты КАК Агенты
        ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.Проектыконтрагентов КАК Проекты
        ПО Агенты.Ссылка = Проекты.Владелец
        ГДЕ Агенты.БИНИИН = "{contragent}"
    """

    query = re.sub(r" {2,}", "", query).strip()

    return query


def iso_to_standard(dt: str) -> str:
    if dt[2] == "." and dt[5] == ".":
        return dt
    return datetime.fromisoformat(dt).strftime("%d.%m.%Y")


@dataclasses.dataclass(slots=True)
class Contract:
    contract_id: str
    contragent: str
    project: str
    credit_purpose: str
    repayment_procedure: str
    loan_amount: float
    subsid_amount: float
    investment_amount: float
    pos_amount: float
    protocol_date: str
    vypiska_date: str
    decision_date: str
    iban: str
    ds_id: str
    ds_date: str
    dbz_id: str
    dbz_date: str
    start_date: str
    end_date: str
    protocol_id: str
    sed_number: str
    document_path: Optional[Path] = None
    protocol_pdf_path: Optional[Path] = None

    def __post_init__(self) -> None:
        self.protocol_date = iso_to_standard(self.protocol_date).replace(".", "")
        self.vypiska_date = iso_to_standard(self.vypiska_date).replace(".", "")
        self.ds_date = iso_to_standard(self.ds_date).replace(".", "")
        self.dbz_date = iso_to_standard(self.dbz_date).replace(".", "")
        self.start_date = iso_to_standard(self.start_date).replace(".", "")
        self.end_date = iso_to_standard(self.end_date).replace(".", "")

        today = str(os.environ["today"])
        contract_folder = ("downloads" / Path(today) / self.contract_id).absolute()
        with suppress(FileNotFoundError):
            self.protocol_pdf_path = next((contract_folder / "vypiska").iterdir(), None)
        self.document_path = contract_folder / "documents" / Path(self.document_path)

    @classmethod
    def iter_contracts(cls, db: DatabaseManager) -> Generator["Contract", None, None]:
        raw_contracts = db.execute("""
            SELECT
                id AS contract_id,
                contragent,
                project,
                credit_purpose,
                repayment_procedure,
                loan_amount,
                subsid_amount,
                investment_amount,
                pos_amount,
                protocol_date,
                vypiska_date,
                decision_date,
                iban,
                ds_id,
                ds_date,
                dbz_id,
                dbz_date,
                start_date,
                end_date,
                protocol_id,
                sed_number,
                file_name
            FROM contracts
            WHERE
                dbz_id IS NOT NULL
                AND id NOT IN (
                    SELECT id FROM errors WHERE traceback IS NOT NULL
                )
        """)

        for raw_contract in raw_contracts:
            contract = Contract(*raw_contract)
            yield contract


@dataclasses.dataclass(slots=True)
class InterestRate:
    contract_id: str
    subsid_term: int
    nominal_rate: float
    rate_one_two_three_year: float
    rate_four_year: float
    rate_five_year: float
    rate_six_seven_year: float
    rate_fee_one_two_three_year: float
    rate_fee_four_year: float
    rate_fee_five_year: float
    rate_fee_six_seven_year: float
    start_date_one_two_three_year: str
    end_date_one_two_three_year: str
    start_date_four_year: str
    end_date_four_year: str
    start_date_five_year: str
    end_date_five_year: str
    start_date_six_seven_year: str
    end_date_six_seven_year: str

    def __post_init__(self) -> None:
        self.nominal_rate *= 100
        self.rate_one_two_three_year *= 100
        self.rate_four_year *= 100
        self.rate_five_year *= 100
        self.rate_six_seven_year *= 100
        self.rate_fee_one_two_three_year *= 100
        self.rate_fee_four_year *= 100
        self.rate_fee_five_year *= 100
        self.rate_fee_six_seven_year *= 100

    @classmethod
    def load(cls, db: DatabaseManager, contract_id: str) -> "InterestRate":
        raw_rate = db.execute(
            """
                SELECT
                    id,
                    subsid_term,
                    nominal_rate,
                    rate_one_two_three_year,
                    rate_four_year,
                    rate_five_year,
                    rate_six_seven_year,
                    rate_fee_one_two_three_year,
                    rate_fee_four_year,
                    rate_fee_five_year,
                    rate_fee_six_seven_year,
                    start_date_one_two_three_year,
                    end_date_one_two_three_year,
                    start_date_four_year,
                    end_date_four_year,
                    start_date_five_year,
                    end_date_five_year,
                    start_date_six_seven_year,
                    end_date_six_seven_year
                FROM interest_rates
                WHERE id = ?
            """,
            (contract_id,),
        )

        return InterestRate(*raw_rate[0])


def find_row(parent: UiaList, project: str) -> Optional[UiaListItem]:
    rows = children(parent)
    for row in rows:
        txt = cast(str, row.window_text()).strip()
        if not txt:
            continue
        score = SequenceMatcher(None, project, txt).ratio()
        if score >= 0.8:
            print(f"{len(rows)=}, {txt=}, {score=}")
            return row
    return None


def find_project(top_win: UiaWindow, contract: Contract) -> None:
    a(top_win, lambda: child_win(top_win, title="Консоль запросов и обработчик", ctrl="Button").click_input())

    query_document_box = child_win(top_win, ctrl="Document")
    query_document_box_obj = query_document_box.wrapper_object()

    delete_button = child_win(top_win, title="Delete", ctrl="Button", idx=1)

    query = prepare_query(contract.contragent)
    pyperclip.copy(query)

    a(top_win, lambda: query_document_box_obj.click_input())
    a(top_win, lambda: send_keys("^a^v"))
    a(top_win, lambda: child_win(top_win, title="Выполнить", ctrl="Button").click_input())

    if not wait_for(lambda: delete_button.is_enabled() == True, timeout=10):
        print(f"{contract.contract_id=} not found")
        return

    row = find_row(child_win(top_win, ctrl="List", idx=1), project=contract.project)
    if not row:
        print(f"{contract.contract_id=} not found")
        return

    a(top_win, lambda: row.double_click_input())

    a(top_win, lambda: query_document_box_obj.click_input())
    query_document_box_obj.type_keys("{ESC}")

    if (close_button := child_win(child_win(top_win, ctrl="Pane", idx=18), title="Close", ctrl="Button")).exists():
        a(top_win, lambda: close_button.click_input())


def fill_main_project_data(win: UiaWindow, form: UiaPane, contract: Contract) -> None:
    """
    :param win: Главное окно 1С
    :param form: Форма "Карточка проекта (форма элемента)"
    :param contract: Данные договора
    :return: None

    Заполнение данных в форме проекта во вкладке "Основные"
    (Цель кредитования, Номер протокола, Дата протокола, Дата получения протокола РКС филиалом)
    """
    a(win, lambda: child_win(parent=form, ctrl="Edit", idx=7).click_input())
    a(win, lambda: send_keys("{F4}^f" + contract.credit_purpose + "{ENTER 2}", pause=0.1))

    a(win, lambda: click_type_keys(child_win(form, ctrl="Edit", idx=1), contract.protocol_id))
    a(win, lambda: click_type_keys(child_win(form, ctrl="Edit", idx=2), contract.protocol_date))
    a(win, lambda: click_type_keys(child_win(form, ctrl="Edit", idx=3), contract.protocol_date))


def change_date(win: UiaWindow, form: UiaPane, goto_button: UiaButton, protocol_date: str) -> None:
    """
    :param win: Главное окно 1С
    :param form: Форма "Карточка проекта (форма элемента)"
    :param goto_button: "Go to" кнопка
    :param protocol_date: Дата протокола
    :return: None

    Возможное изменение даты протокола в форме проекта во вкладке "Пролонгация"
    """
    a(win, lambda: child_win(form, title="Пролонгация", ctrl="TabItem").click_input())

    date_to_check = child_win(form, ctrl="Custom", idx=1).window_text().split(" ")[0].replace(".", "")

    if date_to_check != protocol_date:
        a(win, lambda: goto_button.click_input())
        a(win, lambda: send_keys("{DOWN}{ENTER 2}", pause=0.5))
        a(win, lambda: send_keys(protocol_date))
        a(win, lambda: send_keys("{ENTER 4}{ESC}", pause=0.5))


def change_sums(win: UiaWindow, form: UiaPane, goto_button: UiaButton, contract: Contract) -> None:
    """
    :param win: Главное окно 1С
    :param form: Форма "Карточка проекта (форма элемента)"
    :param goto_button: "Go to" кнопка
    :param contract: Данные договора
    :return: None

    Заполнение данных в форме проекта во вкладке "БВУ/Рефинансирование" в зависимости от цели кредитования
    (Сумма субсидирования, На инвестиции, На ПОС)
    """
    if contract.credit_purpose not in {"Пополнение оборотных средств", "Инвестиционный", "Инвестиционный + ПОС"}:
        raise ValueError(f"Don't know what to do with {contract.credit_purpose!r}...")

    a(win, lambda: child_win(form, title="БВУ/Рефинансирование", ctrl="TabItem").click_input())
    a(win, lambda: goto_button.click_input())
    a(win, lambda: send_keys("{DOWN 8}{ENTER}", pause=0.2))

    list_win = child_win(win, ctrl="Pane", idx=51)

    existing_pos_amount = text_to_float(
        text(child_win(list_win, ctrl="Custom", idx=5)).replace(" Возобновляемая часть", ""), default=0.0
    )
    existing_investment_amount = text_to_float(
        text(child_win(list_win, ctrl="Custom", idx=6)).replace(" Не возобновляемая часть", ""), default=0.0
    )

    a(win, lambda: send_keys("{ENTER}", pause=0.2))

    # record_win = child_win(win, ctrl="Pane", idx=56)

    if contract.credit_purpose == "Пополнение оборотных средств" and existing_pos_amount != contract.subsid_amount:
        a(win, lambda: send_keys("{TAB 4}" + str(contract.subsid_amount), pause=0.1))
    elif contract.credit_purpose == "Инвестиционный" and existing_investment_amount != contract.subsid_amount:
        a(win, lambda: send_keys("{TAB 5}" + str(contract.subsid_amount), pause=0.1))
    elif contract.credit_purpose == "Инвестиционный + ПОС":
        if existing_pos_amount != contract.pos_amount and existing_investment_amount != contract.investment_amount:
            a(win, lambda: send_keys("{TAB 4}" + str(contract.pos_amount) + "{TAB}" + str(contract.investment_amount)))
        elif existing_pos_amount != contract.pos_amount:
            a(win, lambda: send_keys("{TAB 4}" + str(contract.subsid_amount), pause=0.1))
        elif existing_investment_amount != contract.investment_amount:
            a(win, lambda: send_keys("{TAB 5}" + str(contract.subsid_amount), pause=0.1))

    a(win, lambda: send_keys("{ESC}", pause=0.5))
    with suppress(ElementNotFoundError):
        a(win, lambda: child_win(win, title="Yes", ctrl="Button").click_input())
    a(win, lambda: send_keys("{ESC}", pause=0.5))


def add_vypiska(one_c: App, win: UiaWindow, form: UiaPane, contract: Contract) -> None:
    """
    :param one_c: Главный объект
    :param win: Главное окно 1С
    :param form: Форма "Карточка проекта (форма элемента)"
    :param contract: Данные договора
    :return: None

    Прикрепление файла выписки из CRM во вкладке "Прикрепленные документы"
    """

    fname = contract.protocol_pdf_path.name
    file_path = str(contract.protocol_pdf_path)

    a(win, lambda: child_win(form, title="Прикрепленные документы", ctrl="TabItem").click_input())

    a(win, lambda: child_win(form, ctrl="Button", title="Set list filter and sort options...").click_input())

    sort_win = one_c.app.window(title="Filter and Sort")

    check(child_win(sort_win, title="Наименование файла", ctrl="CheckBox"))

    a(win, lambda: click_type_keys(child_win(sort_win, ctrl="Edit", idx=7), fname, spaces=True, escape_chars=True))

    a(win, lambda: child_win(sort_win, title="OK", ctrl="Button").click_input())

    sleep(1)

    if contains_text(child_win(form, ctrl="Table")):
        return

    a(win, lambda: child_win(form, title="Add", ctrl="Button").click_input())
    a(win, lambda: child_win(win, ctrl="Edit", idx=5).click_input())
    a(win, lambda: send_keys("{F4}"))

    one_c.switch_backend("win32")
    save_dialog = one_c.app.window(title="Выберите файл")
    save_dialog["&Имя файла:Edit"].set_text(file_path)
    a(win, lambda: save_dialog.child_win(title="&Открыть", class_name="Button").click_input())
    one_c.switch_backend("uia")

    if (child_win(win, title="Value is not of object type (Сессия)", ctrl="Pane")).exists():
        a(win, lambda: child_win(win, title="OK", ctrl="Button").click_input())
        sleep(1)

    a(win, lambda: child_win(win, title="OK", ctrl="Button", idx=1).click_input())


def check_project_type(win: UiaWindow, form: UiaPane, contract: Contract) -> None:
    if contract.credit_purpose == "Пополнение оборотных средств":
        a(win, lambda: child_win(form, title="Признаки проекта", ctrl="TabItem").click_input())
        check(child_win(form, title="Возобновляемый проект", ctrl="CheckBox"))


def fill_contract_details(win: UiaWindow, ds_form: UiaPane, contract: Contract, rate: InterestRate) -> None:
    # FIXME POS CHANGES INDICES

    # child_window(ds_form, ctrl="Edit", idx=2).set_text(contract.ds_id)
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=20), contract.iban, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=12), contract.ds_id, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=13), contract.ds_date, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=7), contract.dbz_id, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=8), contract.dbz_date, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=9), contract.dbz_date, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=10), contract.end_date, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=11), rate.nominal_rate, ent=True))
    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=14), contract.loan_amount, ent=True))

    if contract.credit_purpose == "Пополнение оборотных средств":
        a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=4), rate.rate_fee_one_two_three_year, ent=True))
        a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=18), contract.pos_amount, ent=True))
    elif contract.credit_purpose == "Инвестиционный":
        a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=19), contract.investment_amount, ent=True))
    elif contract.credit_purpose == "Инвестиционный + ПОС":
        a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=18), contract.pos_amount, ent=True))
        a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=19), contract.investment_amount, ent=True))
    else:
        raise ValueError(f"Don't know what to do with {contract.credit_purpose!r}...")

    a(win, lambda: click_type_keys(child_win(ds_form, ctrl="Edit", idx=3), contract.decision_date, ent=True))

    # Вид погашения платежа - Аннуитетный/Равными долями/Индивидуальный
    a(win, lambda: child_win(ds_form, ctrl="Edit", idx=17).click_input())
    if contract.repayment_procedure == "Аннуитетный":
        a(win, lambda: send_keys("{F4}{ENTER}", pause=0.5))
    elif contract.repayment_procedure == "Равными долями":
        a(win, lambda: send_keys("{F4}{DOWN}{ENTER}", pause=0.5))
    elif contract.repayment_procedure == "Индивидуальный":
        a(win, lambda: send_keys("{F4}{DOWN 2}{ENTER}", pause=0.5))
    else:
        raise ValueError(f"Don't know what to do with {contract.repayment_procedure!r}...")


def fill_contract(one_c: App, win: UiaWindow, form: UiaPane, contract: Contract, rate: InterestRate) -> UiaPane:
    a(win, lambda: child_win(form, title="БВУ/Рефинансирование", ctrl="TabItem").click_input())

    # act(top_win, lambda: child_window(form, ctrl="Custom").click_input())
    # sleep(0.5)
    # act(top_win, lambda: send_keys("{DOWN 10}"))
    # act(top_win, lambda: child_window(form, title="Clone", ctrl="Button", idx=1).click_input())

    a(win, lambda: child_win(form, title="Add", ctrl="Button").click_input())
    ds_form = win.child_window(control_type="Pane", found_index=51)

    fill_contract_details(win, ds_form, contract, rate)

    a(win, lambda: child_win(ds_form, title="Основные реквизиты", ctrl="TabItem").click_input())
    a(win, lambda: child_win(ds_form, title="Распоряжения на изменения статуса договора", ctrl="Button").click_input())

    change_ds_status_form = child_win(win, ctrl="Pane", idx=74)

    table = child_win(change_ds_status_form, ctrl="Table")

    contract_field = child_win(table, ctrl="Custom", idx=1)
    existing_contract = text(contract_field).replace(" Договор субсидирования", "")
    if not existing_contract:
        a(win, lambda: click_type_keys(contract_field, "{F4}", double=True, pause=0.1))
        dict_win = child_win(win, ctrl="Pane", idx=88)
        a(win, lambda: child_win(dict_win, ctrl="Button", title="Set list filter and sort options...").click_input())
        sort_win = one_c.app.window(title="Filter and Sort")

        check(child_win(sort_win, title="Deletion mark", ctrl="CheckBox"))
        check(child_win(sort_win, title="Номер договора субсидирования", ctrl="CheckBox"))

    a(win, lambda: click_type_keys(child_win(table, ctrl="Custom", idx=2), contract.ds_date, double=True, ent=True))
    a(win, lambda: click_type_keys(child_win(table, ctrl="Custom", idx=3), "Подписан ДС", double=True, ent=True, spaces=True))

    for i, cell in enumerate(iter_children(table)):
        txt = text(cell)
        if not txt:
            continue

        value, column = txt.split(" ", maxsplit=1)
        print(i, value, column)

        if column == "Договор субсидирования":
            a(win, lambda: click_type_keys(cell, "{F4}", double=True, pause=0.1))
            dict_win = child_win(win, ctrl="Pane", idx=88)
            a(win, lambda: child_win(dict_win, ctrl="Button", title="Set list filter and sort options...").click_input())
            sort_win = one_c.app.window(title="Filter and Sort")

            check(child_win(sort_win, title="Deletion mark", ctrl="CheckBox"))
            check(child_win(sort_win, title="Номер договора субсидирования", ctrl="CheckBox"))

            # a(
            #     win,
            #     lambda: click_type_keys(
            #         child_win(sort_win, ctrl="Edit", idx=7), fname, spaces=True, escape_chars=True
            #     ),
            # )

            a(win, lambda: click_type_keys(cell, contract.ds_id + "{ENTER 2}", pause=0.1))

        if column == "Дата применения нового статуса":
            a(win, lambda: click_type_keys(cell, contract.ds_date, double=True))

        if "Новый статус" in column:
            a(win, lambda: click_type_keys(cell, "Подписан ДС", double=True, spaces=True))

    a(win, lambda: child_win(change_ds_status_form, title="OK", ctrl="Button").click_input())
    if (close_button := child_win(win, ctrl="Pane", idx=18).child_window(title="Close", control_type="Button")).exists():
        a(win, lambda: close_button.click_input())

    return ds_form


def process_contract(contract: Contract, rate: InterestRate) -> None:
    with App(app_path=r"C:\Users\robot3\Desktop\damu_1c\test_base.v8i") as one_c:
        win = window(one_c.app, title="Конфигурация.+", regex=True)
        win.wait(wait_for="exists", timeout=20)

        find_project(top_win=win, contract=contract)

        form = child_win(win, ctrl="Pane", idx=27)

        fill_main_project_data(win, form, contract)

        goto_button = child_win(form, title="Go to", ctrl="Button")
        change_date(win, form, goto_button, contract.protocol_date)
        change_sums(win, form, goto_button, contract)
        add_vypiska(one_c, win, form, contract)
        check_project_type(win, form, contract)

        ds_form = fill_contract(one_c, win, form, contract, rate)

        # act(top_win, lambda: child_window(change_ds_status_form, title="Закрыть", ctrl="Button").click_input())
        # act(top_win, lambda: child_window(top_win, title="No", ctrl="Button").click_input())

        a(win, lambda: child_win(ds_form, title="Записать", ctrl="Button").click_input())

        a(win, lambda: child_win(ds_form, title="ПрикрепленныеДокументы", ctrl="TabItem").click_input())

        a(win, lambda: child_win(ds_form, title="Add", ctrl="Button").click_input())
        sleep(1)
        a(win, lambda: send_keys("{F4}"))

        one_c.switch_backend("win32")
        save_dialog = one_c.app.window(title_re="Выберите ф.+")
        save_dialog["&Имя файла:Edit"].set_text(str(contract.document_path))
        a(win, lambda: save_dialog.child_win(title="&Открыть", class_name="Button").click_input())
        one_c.switch_backend("uia")

        if (child_win(win, title="Value is not of object type (Сессия)", ctrl="Pane")).exists():
            a(win, lambda: child_win(win, title="OK", ctrl="Button").click_input())
            sleep(1)

        a(win, lambda: child_win(win, title="OK", ctrl="Button", idx=2).click_input())

        a(win, lambda: child_win(ds_form, title="Открыть текущий График погашения", ctrl="Button").click_input())
        a(win, lambda: child_win(win, title="Yes", ctrl="Button").click_input())

        sleep(5)

        table_form = child_win(win, ctrl="Pane", idx=63)
        a(win, lambda: child_win(table_form, ctrl="Edit", idx=9).click_input())
        a(win, lambda: send_keys("13{ENTER}"))
        a(win, lambda: child_win(table_form, ctrl="Edit", idx=5).click_input())
        a(win, lambda: send_keys(contract.start_date + "{ENTER}"))
        a(win, lambda: child_win(table_form, ctrl="Edit", idx=6).click_input())
        a(win, lambda: send_keys(contract.end_date + "{ENTER}"))

        a(win, lambda: child_win(table_form, title="Загрузить из внешней таблицы (обн)", ctrl="Button").click_input())

        # r"C:\Users\robot3\Desktop\damu_1c\downloads\2025-02-26\shifted.xlsx"

        one_c.switch_backend("win32")
        save_dialog = one_c.app.window(title_re="Выберите ф.+")
        save_dialog["&Имя файла:Edit"].set_text(r"C:\Users\robot3\Desktop\damu_1c\downloads\2025-02-26\shifted.xlsx")
        a(win, lambda: save_dialog.child_win(title="&Открыть", class_name="Button").click_input())
        one_c.switch_backend("uia")

        if (close_button := child_win(win, ctrl="Pane", idx=18).child_window(title="Close", control_type="Button")).exists():
            a(win, lambda: close_button.click_input())

        a(win, lambda: child_win(table_form, title="Записать", ctrl="Button").click_input())
        a(win, lambda: child_win(table_form, title="Закрыть", ctrl="Button").click_input())

        a(win, lambda: child_win(ds_form, title="Передать на проверку", ctrl="Button").click_input())

        # pass
        #
        # for _ in range(row_count):
        #     act(top_win, lambda: delete_button.click_input())
        #
        # act(top_win, lambda: query_document_box_obj.click_input())
        # query_document_box_obj.type_keys("{ESC}")


def main() -> None:
    logger.setup_logger(project_folder)

    if sys.version_info.major != 3 or sys.version_info.minor != 12:
        error_msg = f"Python {sys.version_info} is not supported"
        logging.error(error_msg)
        raise RuntimeError(error_msg)

    warnings.simplefilter(action="ignore", category=UserWarning)
    warnings.simplefilter(action="ignore", category=InsecureRequestWarning)
    warnings.simplefilter(action="ignore", category=SyntaxWarning)

    today = datetime(2025, 3, 14).date()
    os.environ["today"] = today.isoformat()

    env_path = project_folder / ".env"
    dotenv.load_dotenv(env_path)

    resources_folder = Path("resources")

    database = resources_folder / "database.sqlite"
    with DatabaseManager(database) as db:
        for contract in Contract.iter_contracts(db):
            if contract.project is None:
                continue

            rate = InterestRate.load(db, contract.contract_id)

            process_contract(contract, rate)


if __name__ == "__main__":
    main()
