import dataclasses
import logging
import os
import shutil
from enum import Enum
from pathlib import Path
from typing import Any, Union, List, Iterator

import win32com.client as win32
import win32com

from src.utils.utils import kill_all_processes


class OfficeType(Enum):
    ExcelType: str = "Excel.Application"
    WordType: str = "Word.Application"


class FileFormat(Enum):
    DOCX: int = 16
    PDF: int = 17


class UnsupportedOfficeAppError(Exception):
    def __init__(self, office_type: OfficeType) -> None:
        message = f"Unknown {office_type!r}"
        super().__init__(message)


def validate_format(file_path: str, file_format: FileFormat) -> bool:
    file_extension = file_path.rsplit(".")[-1]

    match file_format:
        case file_format.DOCX:
            return file_extension == "docx"
        case file_format.PDF:
            return file_extension == "pdf"
        case _:
            return False


class Office:
    def __init__(self, file_path: Union[str, Path], office_type: OfficeType) -> None:
        self.office_type = office_type

        self.file_path: str = str(file_path) if isinstance(file_path, Path) else file_path
        self.project_folder = os.getenv("project_folder")
        if self.project_folder:
            self.file_path = os.path.join(self.project_folder, self.file_path)
        try:
            self.app = win32.Dispatch(office_type.value)
        except AttributeError:
            shutil.rmtree(win32com.__gen_path__)
            self.app = win32.Dispatch(office_type.value)

        self.app.Visible = False
        self.app.DisplayAlerts = False

        self.potential_error = UnsupportedOfficeAppError(office_type=office_type)

        match office_type:
            case OfficeType.ExcelType:
                self.doc = self.open_workbook()
            case OfficeType.WordType:
                self.doc = self.open_doc()
            case _:
                raise self.potential_error

    def open_doc(self) -> Any:
        if self.office_type != OfficeType.WordType:
            raise self.potential_error
        return self.app.Documents.Open(self.file_path)

    def open_workbook(self) -> Any:
        if self.office_type != OfficeType.ExcelType:
            raise self.potential_error
        return self.app.Workbooks.Open(self.file_path)

    def save_as(self, file_path: Union[str, Path], file_format: FileFormat) -> None:
        file_path: str = str(file_path) if isinstance(file_path, Path) else file_path
        if not validate_format(file_path=file_path, file_format=file_format):
            raise ValueError(f"File format and extension mismatch - {file_path!r} {file_format!r}")

        if self.project_folder:
            file_path = os.path.join(self.project_folder, file_path)
        self.doc.SaveAs(file_path, FileFormat=file_format.value)

    def close_doc(self) -> None:
        if not self.doc:
            return

        try:
            self.doc.Close()
        except (Exception, BaseException) as err:
            logging.exception(err)
            kill_all_processes(proc_name="WINWORD")

    def quit_app(self) -> None:
        if not self.app:
            return

        try:
            self.app.Quit()
        except (Exception, BaseException) as err:
            logging.exception(err)
            kill_all_processes(proc_name="WINWORD")
        del self.app

    def __enter__(self) -> "Office":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close_doc()
        self.quit_app()


@dataclasses.dataclass(slots=True)
class Message:
    subject: str
    body: str
    to: str
    attachments: List[Path] = dataclasses.field(default_factory=list)


class Outlook:
    def __init__(self) -> None:
        self.app = win32.Dispatch("Outlook.Application")
        self.namespace = self.app.GetNamespace("MAPI")

    def __enter__(self) -> "Outlook":
        return self

    def send(self, message: Message) -> None:
        mail = self.app.CreateItem(0)
        mail.Subject = message.subject
        mail.Body = message.body
        mail.To = message.to
        for attachment in message.attachments:
            mail.Attachments.Add(str(attachment))
        mail.Send()

    def read_inbox(self, folder: str = "Inbox") -> Iterator[Message]:
        inbox = self.namespace.GetDefaultFolder(6)
        if folder.lower() != "inbox":
            inbox = inbox.Folders[folder]
        for item in inbox.Items:
            if item.Class == 43:
                yield Message(
                    subject=item.Subject,
                    body=item.Body,
                    to=item.To,
                    attachments=[Path(a.FileName) for a in item.Attachments],
                )

    def quit_app(self) -> None:
        if not self.app:
            return

        try:
            self.app.Quit()
        except (Exception, BaseException) as err:
            logging.exception(err)
            kill_all_processes(proc_name="OUTLOOK.EXE")
        del self.app

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.quit_app()
