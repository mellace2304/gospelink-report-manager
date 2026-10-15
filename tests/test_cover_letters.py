"""docx_to_pdf: failures are reported, and bad PDFs from an interrupted run are retried."""
import pytest

import Merge

GOOD_PDF = b"%PDF-1.4\n" + b"x" * 200


class FakeWord:
    """Stands in for Word.Application. Docs named in `fail` raise; `empty` ones save a 0-byte PDF."""
    def __init__(self, fail=(), empty=()):
        self.fail, self.empty = set(fail), set(empty)
        self.opened = []
        self.Visible = self.DisplayAlerts = None
        self.Documents = self

    def Open(self, path):
        name = path.replace("\\", "/").rsplit("/", 1)[-1]
        self.opened.append(name)
        if name in self.fail:
            raise RuntimeError("The RPC server is unavailable")
        return FakeDoc(name in self.empty)

    def Quit(self):
        pass


class FakeDoc:
    def __init__(self, empty):
        self.empty = empty

    def SaveAs(self, path, FileFormat):
        with open(path, "wb") as f:
            f.write(b"" if self.empty else GOOD_PDF)

    def Close(self, save):
        pass


@pytest.fixture
def word(monkeypatch):
    def install(**kw):
        fake = FakeWord(**kw)
        monkeypatch.setattr(Merge.win32com.client, "DispatchEx", lambda _name: fake)
        return fake
    return install


def make_docx(folder, *names):
    for n in names:
        (folder / f"{n}.docx").write_bytes(b"docx")


def test_failures_are_returned_and_bad_output_removed(tmp_path, word):
    make_docx(tmp_path, "1001", "1002", "1003")
    word(fail=["1002.docx"], empty=["1003.docx"])
    failures = Merge.docx_to_pdf(tmp_path, tmp_path)
    assert [name for name, _ in failures] == ["1002.docx", "1003.docx"]
    assert "RPC server" in failures[0][1]
    assert (tmp_path / "1001.pdf").exists()
    assert not (tmp_path / "1002.pdf").exists()
    assert not (tmp_path / "1003.pdf").exists()  # zero-byte output must not linger


def test_zero_byte_pdf_from_earlier_run_is_reconverted(tmp_path, word):
    make_docx(tmp_path, "2001")
    (tmp_path / "2001.pdf").write_bytes(b"")
    fake = word()
    assert Merge.docx_to_pdf(tmp_path, tmp_path) == []
    assert fake.opened == ["2001.docx"]
    assert Merge._validate_pdf(str(tmp_path / "2001.pdf"))[0]


def test_valid_pdf_is_not_reconverted(tmp_path, word):
    make_docx(tmp_path, "3001")
    (tmp_path / "3001.pdf").write_bytes(GOOD_PDF)
    fake = word()
    assert Merge.docx_to_pdf(tmp_path, tmp_path) == []
    assert fake.opened == []


def test_word_lock_files_are_ignored(tmp_path, word):
    make_docx(tmp_path, "4001", "~$4001")
    fake = word()
    Merge.docx_to_pdf(tmp_path, tmp_path)
    assert fake.opened == ["4001.docx"]
