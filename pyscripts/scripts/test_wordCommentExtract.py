import unittest
from unittest.mock import patch, MagicMock
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import os

from wordCommentExtract import (
    select_input_files,
    select_output_file,
    clear_input_files,
    start_extraction,
    extract_comments,
    extract_comments_from_docx,
    get_page_number
)

class TestWordCommentExtract(unittest.TestCase):

    def setUp(self):
        self.root = tk.Tk()
        self.input_files_list = tk.Text(self.root)
        self.output_file_var = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.root.update_idletasks()

    @patch('wordCommentExtract.filedialog.askopenfilenames')
    def test_select_input_files(self, mock_askopenfilenames):
        mock_askopenfilenames.return_value = ['/path/to/file1.docx', '/path/to/file2.docx']
        select_input_files()
        self.assertEqual(self.input_files_list.get(1.0, tk.END).strip(), '/path/to/file1.docx\n/path/to/file2.docx')

    @patch('wordCommentExtract.filedialog.asksaveasfilename')
    def test_select_output_file(self, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = '/path/to/output.xlsx'
        select_output_file()
        self.assertEqual(self.output_file_var.get(), '/path/to/output.xlsx')

    def test_clear_input_files(self):
        self.input_files_list.insert(tk.END, 'Some text')
        clear_input_files()
        self.assertEqual(self.input_files_list.get(1.0, tk.END).strip(), '')

    @patch('wordCommentExtract.threading.Thread')
    def test_start_extraction(self, mock_thread):
        start_extraction()
        mock_thread.assert_called_once()

    @patch('wordCommentExtract.messagebox.showerror')
    @patch('wordCommentExtract.openpyxl.Workbook')
    @patch('wordCommentExtract.extract_comments_from_docx')
    @patch('wordCommentExtract.os.path.exists')
    def test_extract_comments(self, mock_exists, mock_extract_comments_from_docx, mock_Workbook, mock_showerror):
        mock_exists.return_value = True
        mock_extract_comments_from_docx.return_value = [{'author': 'Author', 'text': 'Comment', 'page': 1, 'referenced_text': 'Text'}]
        mock_ws = MagicMock()
        mock_wb = MagicMock()
        mock_wb.active = mock_ws
        mock_Workbook.return_value = mock_wb

        self.input_files_list.insert(tk.END, '/path/to/file1.docx\n/path/to/file2.docx')
        self.output_file_var.set('/path/to/output.xlsx')

        extract_comments()

        mock_ws.append.assert_any_call(["File Name", "Comment Author", "Comment Text", "Page Number", "Referenced Text"])
        mock_ws.append.assert_any_call(['file1.docx', 'Author', 'Comment', 1, 'Text'])
        mock_wb.save.assert_called_once_with('/path/to/output.xlsx')
        mock_showerror.assert_not_called()

    @patch('wordCommentExtract.zipfile.ZipFile')
    @patch('wordCommentExtract.docx.Document')
    def test_extract_comments_from_docx(self, mock_Document, mock_ZipFile):
        mock_doc = MagicMock()
        mock_Document.return_value = mock_doc
        mock_doc.paragraphs = [MagicMock(text='Referenced text', runs=[MagicMock(_element=MagicMock(find=MagicMock(return_value=None)))]),]
        mock_zip = MagicMock()
        mock_ZipFile.return_value.__enter__.return_value = mock_zip
        mock_zip.read.return_value = b'<comments xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><comment id="0" author="Author"><p><t>Comment</t></p></comment></comments>'

        comments = extract_comments_from_docx('/path/to/file.docx')
        self.assertEqual(len(comments), 0)

    def test_get_page_number(self):
        mock_doc = MagicMock()
        mock_doc.paragraphs = [MagicMock(runs=[MagicMock(element=MagicMock(tag='w:br', get=MagicMock(return_value='page')))]),]
        page_number = get_page_number(mock_doc, 0)
        self.assertEqual(page_number, 1)

if __name__ == '__main__':
    unittest.main()