from docxtpl import DocxTemplate
import datetime
import re
from pathlib import Path
import os
import argparse
import sys
import docx
import asciimatics
from asciimatics.widgets import (
    Frame,
    ListBox,
    Layout,
    Divider,
    Text,
    Button,
    TextBox,
    Widget,
    DatePicker,
    DropdownList,
    MultiColumnListBox,
    PopUpDialog,
)
from asciimatics.event import KeyboardEvent
from asciimatics.scene import Scene
from asciimatics.screen import Screen
from asciimatics.exceptions import ResizeScreenError, NextScene, StopApplication
import sys
import sqlite3


class EscapeFrame(Frame):
    def process_event(self, event):
        if isinstance(event, KeyboardEvent) and event.key_code == Screen.KEY_ESCAPE:
            self.cancel()
        else:
            super().process_event(event)

    def cancel(self):
        # do nothing by default
        pass


class ValidationError(Exception):
    def __init__(self, msg):
        Exception.__init__(self, msg)


def adapt_date_iso(val):
    """Adapt datetime.date to ISO 8601 date."""
    return val.isoformat()


def convert_date(val):
    """Convert ISO 8601 date to datetime.date object."""
    return datetime.date.fromisoformat(val.decode())


sqlite3.register_adapter(datetime.date, adapt_date_iso)
sqlite3.register_converter("date", convert_date)


class RefLetterModel:
    def __init__(self):
        # Create a database in RAM.
        self._db = sqlite3.connect(
            "reference_list.sqlite3", detect_types=sqlite3.PARSE_DECLTYPES
        )
        self._db.row_factory = sqlite3.Row

        # Create the basic refletter table.
        self._db.cursor().execute(
            """
            CREATE TABLE if not exists refletters(
                id INTEGER PRIMARY KEY,
                name TEXT,
                ref_date DATE,
                start_year TEXT,
                end_year TEXT,
                target TEXT,
                how_known TEXT,
                recommendation TEXT)
        """
        )
        self._db.commit()

        # Current refletter when editing.
        self.current_id = None

    def add(self, refletter):
        cursor = self._db.cursor()
        cursor.execute(
            """
            INSERT INTO refletters(name, ref_date, start_year, end_year, how_known,recommendation)
            VALUES(:name, :ref_date, :start_year, :end_year, :how_known, :recommendation)""",
            refletter,
        )
        row_id = cursor.lastrowid
        self._db.commit()
        new_id = (
            self._db.cursor()
            .execute("SELECT id from refletters where rowid=?", [row_id])
            .fetchone()["id"]
        )
        return new_id

    def duplicate(self):
        if self.current_id is not None:
            data = dict(self.get_current_refletter())
            del data["id"]
            data["ref_date"] = datetime.date.today()
            self.current_id = self.add(data)

    def get_summary(self):
        def str_not_none(s):
            if s == None:
                return ""
            else:
                return str(s)

        summary_lines = (
            self._db.cursor()
            .execute(
                "SELECT  ref_date,name,target,id from refletters order by ref_date desc"
            )
            .fetchall()
        )
        return [
            ((str(d), str_not_none(n), str_not_none(t)), record_id)
            for d, n, t, record_id in summary_lines
        ]

    def get_refletter(self, refletter_id):
        return (
            self._db.cursor()
            .execute(
                "SELECT id, * from refletters WHERE id=:id",
                {"id": refletter_id},
            )
            .fetchone()
        )

    def get_current_refletter(self):
        if self.current_id is None:
            return {
                "ref_date": datetime.date.today(),
                "target": "",
                "name": "",
                "start_year": "",
                "end_year": "",
                "how_known": "",
                "recommendation": "",
            }
        else:
            details = dict(self.get_refletter(self.current_id))
            if details["end_year"] is None:
                details["end_year"] = ""
            if type(details["start_year"]) == str:
                details["start_year"] = int(details["start_year"])

            return details

    def validate_refletter(self, details):
        def check_val(details, val):
            if val not in details:
                return True
            if details[val] is None:
                return True
            if len(str(details[val])) == 0:
                return True
            return False

        if check_val(details, "start_year"):
            raise ValidationError(f"Start year missing{details}")
        elif check_val(details, "recommendation"):
            raise ValidationError(
                f"Missing recommendation:{details} {list(details.keys())}"
            )
        elif check_val(details, "how_known"):
            raise ValidationError("Missing how you know them")

    def update_current_refletter(self, details):
        self.validate_refletter(details)
        if self.current_id is None:
            self.add(details)
        else:
            columns = ",".join(["%s=:%s" % (col, col) for col in details.keys()])
            self._db.cursor().execute(
                f"""
                UPDATE refletters SET {columns} WHERE id=:id""",
                details,
            )
            self._db.commit()

    def delete_refletter(self, refletter_id):
        self._db.cursor().execute(
            """
            DELETE FROM refletters WHERE id=:id""",
            {"id": refletter_id},
        )
        self._db.commit()

    def write_docx(self):
        if self.current_id is not None:
            data = dict(self.get_current_refletter())
            self.validate_refletter(data)
            date_text = data["ref_date"].strftime("%d %b %Y")

            data["has_end"] = len(data["end_year"].strip()) != 0
            context = {
                "ref_date": date_text,
                "start_date": "Sep %s" % data["start_year"],
                "has_end": data["has_end"],
                "how_known": data["how_known"],
                "end_date": "June %s" % data["end_year"],
                "recommendation_text": data["recommendation"],
                "student_name": data["name"],
                "target": data["target"],
            }
            name_clean = re.sub(r"\W+", "_", data["name"])
            filename = Path(
                datetime.date.today().strftime(f"%Y-%m-%d-{name_clean}.docx")
            )
            doc = DocxTemplate("reference_template.docx")
            try:
                doc.render(context)
            except docx.opc.exceptions.PackageNotFoundError:
                return False
            doc.save(filename)
            pdf_name = str(filename.absolute().with_suffix(".pdf"))
            if sys.platform == "win32":
                import comtypes
                import comtypes.client

                word = comtypes.client.CreateObject("Word.Application")
                word.Visible = 1
                doc = word.Documents.Open(str(filename.absolute()))
                doc.SaveAs(pdf_name, 17)
                doc.Close()
                os.startfile(pdf_name)
            else:
                # TODO: other OS support for
                # docx -> pdf conversion
                os.startfile(filename)
            return True


class LetterList(EscapeFrame):
    def __init__(self, screen, model):
        super(LetterList, self).__init__(
            screen,
            screen.height,
            screen.width,
            on_load=self._reload_list,
            hover_focus=True,
            can_scroll=False,
            title="List of references",
        )
        # Save off the model that accesses the contacts database.
        self._model = model
        self.palette = custom_colour_theme

        # Create the form for displaying the list of contacts.
        self._list_view = MultiColumnListBox(
            Widget.FILL_FRAME,
            columns=["<15%", "<40%", "<40%"],
            options=model.get_summary(),
            name="letters",
            add_scroll_bar=True,
            on_change=self._on_pick,
            on_select=self._edit,
        )
        self._generate_button = Button("Generate (F5)", self._generate)
        self._copy_button = Button("Copy", self._copy)
        self._edit_button = Button("Edit", self._edit)
        self._delete_button = Button("Delete", self._delete)
        layout = Layout([100], fill_frame=True)
        self.add_layout(layout)
        layout.add_widget(self._list_view)
        layout.add_widget(Divider())
        layout2 = Layout([1, 1, 1, 1, 1, 1])
        self.add_layout(layout2)
        layout2.add_widget(Button("Add", self._add), 0)
        layout2.add_widget(self._copy_button, 1)
        layout2.add_widget(self._edit_button, 2)
        layout2.add_widget(self._delete_button, 3)
        layout2.add_widget(self._generate_button, 4)
        layout2.add_widget(Button("Quit", self._quit), 5)
        self.fix()
        self._on_pick()

    def _on_pick(self):
        self._edit_button.disabled = self._list_view.value is None
        self._delete_button.disabled = self._list_view.value is None
        self._copy_button.disabled = self._list_view.value is None

    def _reload_list(self, new_value=None):
        self._list_view.options = self._model.get_summary()
        self._list_view.value = new_value

    def reset(self):
        super(LetterList, self).reset()
        if self._model.current_id is not None:
            self._list_view.value = self._model.current_id

    def process_event(self, event):
        if isinstance(event, KeyboardEvent) and event.key_code == Screen.KEY_F5:
            self._generate()
        else:
            super().process_event(event)

    def _add(self):
        self._model.current_id = None
        raise NextScene("Edit Reference")

    def _copy(self):
        self.save()
        self._model.current_id = self.data["letters"]
        self._model.duplicate()
        raise NextScene("Edit Reference")

    def _generate(self):
        self.save()
        self._model.current_id = self.data["letters"]
        if self._model.write_docx() == False:
            self._scene.add_effect(
                PopUpDialog(
                    self.screen,
                    "Couldn't write docx file - is the template open and locked in word?",
                    buttons=["OK"],
                )
            )

    def _edit(self):
        self.save()
        self._model.current_id = self.data["letters"]
        raise NextScene("Edit Reference")

    def _delete(self):
        self.save()
        self._model.delete_contact(self.data["letters"])
        self._reload_list()

    def cancel(self):
        self._quit()

    @staticmethod
    def _quit():
        raise StopApplication("User pressed quit")


class LetterView(EscapeFrame):
    def __init__(self, screen, model):
        super(LetterView, self).__init__(
            screen,
            screen.height * 2 // 3,
            screen.width * 2 // 3,
            hover_focus=True,
            can_scroll=False,
            title="Reference Details",
            reduce_cpu=True,
        )
        # Save off the model that accesses the contacts database.
        self._model = model

        # Create the form for displaying the list of contacts.
        layout = Layout([100], fill_frame=True)
        self.add_layout(layout)
        layout.add_widget(Text("Name:", "name"))
        layout.add_widget(DatePicker("Reference date:", "ref_date"))
        year_range = range(
            datetime.date.today().year - 10, datetime.date.today().year + 1
        )
        year_range = list(reversed(year_range))
        layout.add_widget(
            DropdownList(
                options=[(str(x), x) for x in year_range],
                label="Start year:",
                name="start_year",
            )
        )
        layout.add_widget(
            DropdownList(
                options=[("Still here", "")] + [(str(x), x) for x in year_range],
                label="End year:",
                name="end_year",
            )
        )
        layout.add_widget(Text("Target course or university:", "target"))
        layout.add_widget(Text("How known:", "how_known"))
        layout.add_widget(
            TextBox(
                Widget.FILL_FRAME,
                "Recommendation:",
                "recommendation",
                as_string=True,
                line_wrap=True,
            )
        )
        layout2 = Layout([1, 1, 1, 1])
        self.add_layout(layout2)
        layout2.add_widget(Button("OK", self._ok), 0)
        layout2.add_widget(Button("Cancel", self.cancel), 3)
        self.fix()

    def reset(self):
        # Do standard reset to clear out form, then populate with new data.
        super(LetterView, self).reset()
        self.data = self._model.get_current_refletter()

    def _strip_full_stops(self, fields):
        for f in fields:
            self.data[f] = self.data[f].strip()
            if self.data[f].endswith("."):
                self.data[f] = self.data[f][:-1]

    def _ok(self):
        self.save()
        self._strip_full_stops(["target", "how_known", "recommendation"])
        try:
            self._model.update_current_refletter(self.data)
            raise NextScene("Main")
        except ValidationError as e:
            message = e.args[0]
            self._scene.add_effect(
                PopUpDialog(self.screen, f"{message}: please fix", buttons=["OK"])
            )
            raise e

    def is_dirty(self):
        self.save()
        return (self._model.get_current_refletter()!=self.data)

    def cancel(self):
        if self.is_dirty():
            def quit_if_1(v):
                if v==1:
                    raise NextScene("Main")
            self._scene.add_effect(
                PopUpDialog(
                    self.screen,
                    "You've made changes, dump them?",
                    buttons=["Cancel","OK"],on_close=quit_if_1
                ))
        else:
            raise NextScene("Main")


custom_colour_theme = dict(asciimatics.widgets.utilities.THEMES["default"])
custom_colour_theme["selected_field"] = (
    Screen.COLOUR_BLACK,
    Screen.A_NORMAL,
    Screen.COLOUR_CYAN,
)


def run_scenes(screen, scene, letters):
    scenes = [
        Scene([LetterList(screen, letters)], -1, name="Main"),
        Scene([LetterView(screen, letters)], -1, name="Edit Reference"),
    ]

    screen.play(scenes, stop_on_resize=True, start_scene=scene, allow_int=True)


def run():
    letters = RefLetterModel()
    last_scene = None
    while True:
        try:
            Screen.wrapper(
                run_scenes, catch_interrupt=True, arguments=[last_scene, letters]
            )
            sys.exit(0)
        except ResizeScreenError as e:
            last_scene = e.scene
