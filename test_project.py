import os
import pandas as pd
import pytest
from project import pull_file_info, write_to_excel, find_new_docs, write_new_docs

# Create a temporary directory and files for testing
@pytest.fixture
def temp_dir(tmp_path):
    temp_dir = tmp_path / "test_folder"
    temp_dir.mkdir()
    (temp_dir / "test1.txt").write_text("This is a test file.")
    (temp_dir / "test2.pdf").write_text("This is another test file.")
    yield temp_dir

def test_pull_file_info(temp_dir):
    result = pull_file_info(temp_dir)
    answer = [
        {
            "File Path": str(temp_dir / "test1.txt"),
            "File Name": "test1",
            "File Type": ".txt",
            "File Size (KB)": str(os.path.getsize(temp_dir / "test1.txt") / (10**3))
        },
        {
            "File Path": str(temp_dir / "test2.pdf"),
            "File Name": "test2",
            "File Type": ".pdf",
            "File Size (KB)": str(os.path.getsize(temp_dir / "test2.pdf") / (10**3))
        }
    ]
    assert result == answer

def test_write_to_excel(temp_dir):
    info = pull_file_info(temp_dir)
    out_folder = temp_dir
    write_to_excel(out_folder, info)
    
    # Check if the Excel file is created
    excel_files = list(temp_dir.glob("*.xlsx"))
    assert len(excel_files) == 1

def test_find_new_docs(temp_dir, tmp_path):
    info = pull_file_info(temp_dir)
    out_folder = tmp_path / "output"
    out_folder.mkdir()
    write_to_excel(out_folder, info)

    # Add a new file ot the temp dir
    (temp_dir / "test3.docx").write_text("This is a new test file.")

    excel_path = list(out_folder.glob("*.xlsx"))[0]
    new_docs_df = find_new_docs(temp_dir, excel_path)

    # Check if the new document is in the df
    assert not new_docs_df.empty
    assert new_docs_df["File Name"].iloc[0] == "test3"

def test_write_new_docs(temp_dir, tmp_path):
    info = pull_file_info(temp_dir)
    out_folder = tmp_path / "output"
    out_folder.mkdir()
    write_to_excel(out_folder, info)

    # Add a new file to the temp dir
    (temp_dir / "test3.docx").write_text("This is a new test file.")
    excel_path = list(out_folder.glob("*.xlsx"))[0]
    new_docs_df = find_new_docs(temp_dir, excel_path)

    # Write the new documents to Excel
    write_new_docs(out_folder, new_docs_df)

    # Check if the new Excel file is created
    new_excel_files = list(out_folder.glob("Added Docs_*.xlsx"))
    assert len(new_excel_files) == 1

    # Check that the new file is in the excel doc
    new_df = pd.read_excel(new_excel_files[0])
    assert "test3" in new_df["File Name"].values

# Made with the generous and eternally patient helpers Copilot, Claude, and Gemini
