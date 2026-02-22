'''
Batch Assignment File Renaming Tool
@author: Leo (geneleo@qq.com)
@version: v1.3.12
@date: 2023.03.24
@environment: Python 3.11+, Windows 10/11, Microsoft Excel
@dependencies: xlwings, opencc
'''

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext
import xlwings as xw
from opencc import OpenCC
import os


class Common:
    """Common constants class"""
    
    ABOUT = '''\
    Main Features: 
    1. Assignment file renaming
    2. Track submission status and save to Excel
    
    Author: Leo (geneleo@qq.com)
    Version: v1.3.12 (2023.03.24)\
    '''
    
    NOTICE = '''\
Notice:

1. Filename must contain the correct student name, otherwise it cannot be renamed
2. Files from students with duplicate names must contain correct student ID for distinction
3. Only processes first-level files/folders in the specified directory, no recursive processing
4. For Excel roster format, please refer to README.md\
    '''
    
    ENTRY_WIDTH = 70


class Rename:
    """Core renaming class"""
    
    excelworkno = 0  # Excel column counter for generating 'Assignment n' headers

    def __init__(self, dirpath, excelfile, samplename, sheetname, *columnname):
        self.dirpath = dirpath
        self.excelfile = excelfile
        self.sheetname = sheetname
        self.samplename = samplename
        self.columnname = columnname
        
        # Excel cell range definitions
        self.column_stuno = "A2:A10000"
        self.column_stuname = "B2:B10000"
        self.column_classname = "C2:C10000"
        self.column_result = "D1"
        
        # Data storage
        self.stuname_extra = ".International"
        self.errorinfo = ''
        self.successinfo = ''
        self.stu = {}  # Student dict {student_id: [name, suffix, class]}
        self.stuno = []
        self.stuname = []
        self.stunameext = []
        self.classname = []
        self.dupSubmitStu = []  # Student IDs who submitted duplicates
        self.submitStu = []  # Student IDs who submitted assignments
        self.studirs = []  # Student IDs who submitted directories
        
        self.start()

    def skipfile(self, cc, sno, fname):
        """Determine if current file should be skipped"""
        sname = self.stu[sno][0]

        if fname.startswith('~'):  # Skip temporary files
            return True
            
        convert_name = cc.convert(fname)  # Convert to simplified Chinese for comparison
        if convert_name.find(sname) == -1:  # Name doesn't match
            return True

        # Students with duplicate names must match student ID
        if self.stuname.count(sname) > 1:
            if convert_name.find(str(sno)) == -1:
                return True
        return False

    def start(self):
        """Main renaming workflow"""
        if self.isDirpathError():
            return

        with xw.App(visible=False) as app:
            book = self.openExcel()
            if not book:
                return

            sheet = self.openSheet(book)
            if not sheet:
                return

            print("\n" + '+' * 55)
            
            # Read Excel data
            if not self.readStuno(sheet) or not self.readStuname(sheet):
                return

            if self.samplename['name']['classname'] and not self.readClassname(sheet):
                return

            self.chineseT2S()
            self.handleStunameExtra()
            self.createStuDict()

            if not self.checkFiles():
                return

            if not self.errorinfo:
                if not self.writeResult(sheet):
                    return

            if not self.saveExcel(book):
                return

        self.rename()

    def writeResult(self, sheet):
        """Write assignment submission status to Excel"""
        print('+ Saving submission records to Excel (not submitted: 0, submitted: 1)')
        startunit = self.column_result
        value = []

        Rename.excelworkno += 1
        no = Rename.excelworkno
        
        # Find empty column
        for i in range(ord('Z') - ord(startunit[0]) + 1):
            col = chr(ord(startunit[0]) + i)
            try:
                vlist = sheet.range(f'{col}1:{col}10000').value
            except Exception as err:
                self.errorinfo = "Error reading Excel data: Please check and retry!"
                return False

            vlist = [str(v).strip() for v in vlist if v is not None]
            if not any(vlist):
                value.append(f'Assignment{no}')
                startunit = f'{col}1'
                break

        for sno in self.stuno:
            v = 1 if sno in self.submitStu else 0
            value.append(v)

        try:
            sheet.range(startunit).options(transpose=True).value = value
        except:
            self.errorinfo = "Error writing to Excel: Check if Excel is already open, please close it and retry!"
            return False
        return True

    def isDirpathError(self):
        """Check if directory is valid"""
        try:
            os.scandir(self.dirpath)
        except:
            self.errorinfo = "Directory error: Specified directory does not exist or has other issues, please correct and retry!"
            return True
        return False

    def checkFiles(self):
        """Check file status and track submission"""
        print('+ Checking files to be renamed')
        submitCount = {}
        cc = OpenCC('t2s')
        try:
            for sno in self.stuno:
                with os.scandir(self.dirpath) as item:  # Open directory as item
                    for entry in item:  # Get each entry (i.e., file) in item
                        if self.skipfile(cc, sno, entry.name):
                            continue
                        if sno not in self.submitStu:
                            self.submitStu.append(sno)  # Add to submitted list
                            submitCount[sno] = 1
                        else:
                            submitCount[sno] += 1

                        # Students who submitted directories
                        if not entry.is_file() and (sno not in self.studirs):
                            self.studirs.append(sno)
                        os.rename(entry.path, entry.path)  # Try renaming to check if file allows rename
                        # Note: Cannot build new name here directly, as duplicate submissions aren't fully counted yet
        except PermissionError as err:
            self.errorinfo = f'  File: "{entry.name}" is already open, please close it and re-run the program!'
            return False
        else:
            # Count students who submitted duplicate files
            for key in submitCount:
                if submitCount[key] >= 2:
                    self.dupSubmitStu.append(key)
            # If there are students who submitted duplicates
            if self.dupSubmitStu:
                namelist = self.getNoNameList(self.dupSubmitStu)  # Get "student_id.name.suffix" list by student ID
                nl = '\n'
                nl2 = '\n    '
                self.errorinfo = f"Students who submitted duplicate assignments (total {len(namelist)} people):{nl}    {nl2.join(namelist)}{nl}"\
                                 f"Duplicate assignments not renamed, please organize and retry"
        return True

    def getNameList(self, stunoList):
        """Generate 'student_id.name.suffix' format list from student IDs"""
        namelist = []
        for i in stunoList:
            namelist.append(f"{i}.{self.stu[i][0]}{self.stu[i][1]}")
        return namelist

    def getNoNameList(self, stunoList):
        """Generate 'sequence_number-student_id.name.suffix' format list from student IDs"""
        namelist = []
        for n, i in enumerate(stunoList):
            namelist.append(f"{n+1:04}-{i}.{self.stu[i][0]}{self.stu[i][1]}")
        return namelist

    def openExcel(self):
        """Open Excel workbook"""
        try:
            book = xw.Book(self.excelfile)
        except:
            self.errorinfo = "File error: Specified Excel file does not exist or Excel is already open, please check and retry!"
            return
        return book

    def openSheet(self, book):
        """Get Excel worksheet"""
        try:
            if self.sheetname:
                sheet = book.sheets(self.sheetname)
            else:
                sheet = book.sheets[0]
        except Exception as err:
            self.errorinfo = 'Error: Opening Excel worksheet failed, check the sheet name?'
            return
        return sheet

    def saveExcel(self, book):
        """Save Excel workbook"""
        try:
            book.save(self.excelfile)
        except Exception as err:
            self.errorinfo = "Error saving Excel data: Is the specified Excel file read-only or already open? Please check and retry!"
            return False
        return True

    def readStuno(self, sheet):
        """Read student ID column from Excel"""
        print('+ Reading student ID column')
        stuno = sheet.range(self.column_stuno).value
        stuno = [x for x in stuno if x is not None]

        if len(stuno) > len(set(stuno)):
            self.errorinfo = "Error: Student IDs in Excel sheet are duplicated (or student ID column data format is incorrect), please correct and retry!"
            return False

        try:                    
            stuno = [int(l) for l in stuno]
        except ValueError:
            self.errorinfo = "Error: Student ID data type in Excel sheet is incorrect, please correct and retry!"
            return False

        self.stuno = stuno
        return True

    def readStuname(self, sheet):
        """Read student name column from Excel"""
        print('+ Reading student name column')
        stuname = sheet.range(self.column_stuname).value
        stuname = [x for x in stuname if x is not None]

        if len(self.stuno) != len(stuname):
            self.errorinfo = "Data error: Student ID column and name column in Excel sheet have inconsistent quantities, please correct and retry."
            return False

        self.stuname = stuname
        return True

    def readClassname(self, sheet):
        """Read class column from Excel"""
        print('+ Reading class column')
        clsname = sheet.range(self.column_classname).value
        clsname = [x for x in clsname if x is not None]

        if not clsname:
            self.errorinfo = "Data error: Class/major column in Excel sheet is empty, please correct and retry."
            return False

        if len(self.stuno) != len(clsname):
            self.errorinfo = "Data error: Number of elements in class/major column and corresponding student IDs in Excel sheet are inconsistent, please correct and retry."
            return False

        self.classname = clsname
        return True

    def chineseT2S(self):
        """Convert traditional Chinese names to simplified"""
        print('+ Converting names to simplified Chinese')
        cc = OpenCC('t2s')
        for i in range(len(self.stuname)):
            self.stuname[i] = cc.convert(self.stuname[i])

    def handleStunameExtra(self):
        """Handle special markers and suffixes in names"""
        print('+ Removing prefixes/suffixes like "*, (online), (makeup)" etc from names')
        
        psFix = [
            '*', '(跟班重修)', '(线上)', '(先修)', '(补修)', 
            '(重修)', '(辅修)', '(（跟班重修）)', '(（线上）)', 
            '(（先修）)', '(（补修）)', '(（重修）)', '(（辅修）)'
        ]
        
        # Remove suffixes, up to 3 times (handle multiple suffixes)
        for i, name in enumerate(self.stuname):
            for _ in range(3):
                for fix in psFix:
                    if name.endswith(fix):
                        name = name[:-len(fix)]
                        self.stuname[i] = name
                    if name.startswith(fix):
                        name = name[len(fix):]
                        self.stuname[i] = name
        
        # Initialize name suffixes
        dupname = {}
        for i, v in enumerate(self.stuname):
            self.stunameext.append('')

        # Handle students with duplicate names
        hasDupStu = False
        for i, v in enumerate(self.stuname):
            if self.stuname.count(v) > 1:
                hasDupStu = True
                dupname[v] = dupname.get(v, 0) + 1
                self.stunameext[i] = str(dupname[v]) + self.stunameext[i]
                
        if hasDupStu:
            print("+  Handling students with duplicate names in class")

    def createStuDict(self):
        """Build student dictionary"""
        print('+ Building student dictionary')
        self.stu = dict.fromkeys(self.stuno)

        for key in self.stu:
            self.stu[key] = ['', '', '']

        for i, name in enumerate(self.stuname):
            self.stu[self.stuno[i]] = [name]

        for i, ext in enumerate(self.stunameext):
            self.stu[self.stuno[i]].append(ext)

        if self.classname:
            for i, cls in enumerate(self.classname):
                self.stu[self.stuno[i]].append(cls)

    def getErrorInfo(self):
        return self.errorinfo

    def getSuccessInfo(self):
        return self.successinfo

    def rename(self):
        """Execute file renaming"""
        print('+ Starting renaming >>>>>>>>>>>>>>>>>>>>>>>>>>>')
        
        serialno = 1
        newname = {}
        
        for v in self.samplename['order']:
            newname[v] = ''
            
        newname['prefix1'] = self.samplename['name']['prefix1']
        newname['prefix2'] = self.samplename['name']['prefix2']
        newname['prefix3'] = self.samplename['name']['prefix3']
        newname['suffix'] = self.samplename['name']['suffix']

        if not self.dirpath.endswith('/'):
            self.dirpath += "/"

        cc = OpenCC('t2s')
        
        for sno in self.submitStu:
            sname = self.stu[sno][0]
            with os.scandir(self.dirpath) as item:
                for entry in item:
                    if self.skipfile(cc, sno, entry.name):
                        continue
                    if sno in self.dupSubmitStu:
                        continue

                    if self.samplename['name']['serialno']:
                        newname['serialno'] = f'{serialno:03d}'
                        serialno += 1
                        
                    if self.samplename['name']['studentno']:
                        newname['studentno'] = str(sno)
                        
                    if self.samplename['name']['studentname']:
                        newname['studentname'] = sname
                        
                    if self.samplename['name']['classname']:
                        newname['classname'] = self.stu[sno][2]
                        
                    if sno in self.studirs:
                        newname['extension'] = ''
                    elif not self.samplename['name']['extension']:
                        newname['extension'] = self.getExt(entry.name)
                    else:
                        newname['extension'] = self.samplename['name']['extension']
                        
                    os.rename(entry.path, self.dirpath + ''.join(newname.values()))
        # Renaming successful
        if not self.errorinfo:
            self.successinfo = ('*' * 24 + ' Renaming Successful! ' + '*' * 20)
            if len(self.submitStu) == len(self.stuno):
                self.successinfo += f'\n    All {len(self.stuno)} assignments submitted!'
            else:
                self.successinfo += ('\n    ' + f'Submitted assignments: {len(self.submitStu)}/{len(self.stuno)}'
                                     + '\n    ' + f'Not submitted: {len(self.stuno) - len(self.submitStu)}, list of students who did not submit:'
                                     + '\n    ' + '--' * 3 + '\n' + '    '
                                     + '\n    '.join(self.getNoNameList(list(set(self.stuno) - set(self.submitStu))))
                                     )
            self.successinfo += '\n' + '*' * 55

    def getExt(self, file):
        """Get file extension (including dot, e.g., .docx)"""
        return os.path.splitext(file)[-1]


class MsgDialog(tk.Toplevel):
    """Message dialog class"""
    
    dialog = {}

    def __init__(self, parent, id, title, icon, message, scrolledtext=False, width=10):
        self.id = id
        if id in MsgDialog.dialog:
            MsgDialog.dialog[id].focus()
            return
            
        super().__init__(parent)
        MsgDialog.dialog[id] = self
        self.protocol("WM_DELETE_WINDOW", self.close)

        MsgDialog.scrolledtext = scrolledtext
        try:
            self.iconbitmap(icon)
        except:
            pass
            
        self.title(title)
        self.transient(parent)
        self.resizable(0, 0)

        if not scrolledtext:
            width = parent.winfo_width() // 2 if width == 10 else width
            msg = tk.Message(self, width=width, text=message, justify="left")
            msg.pack(padx=10, pady=10)
        else:
            self.createScrolledText(message)

    def rgb(self, r, g, b):
        """Generate RGB color string"""
        return "#{:2x}{:2x}{:2x}".format(r, g, b)

    def createScrolledText(self, message):
        """Create scrolled text widget"""
        MsgDialog.text_area = scrolledtext.ScrolledText(
            self, wrap=tk.WORD, width=100, height=30, 
            font=("Times New Roman", 12),
            bg=self.rgb(245, 245, 245), 
            fg=self.rgb(51, 51, 51)
        )
        MsgDialog.text_area.grid(column=0, pady=10, padx=10)
        MsgDialog.text_area.focus()
        MsgDialog.text_area.insert(tk.INSERT, message)

    def close(self):
        """Close dialog"""
        del MsgDialog.dialog[self.id]
        self.destroy()

    @staticmethod
    def updateMessage(msg):
        if not MsgDialog.scrolledtext:
            return
        MsgDialog.text_area.insert(tk.INSERT, msg + '\n')


class Window(tk.Tk):
    """Main window class"""
    
    def __init__(self, title, icon):
        super().__init__()
        self.initVar()
        self.initWindow(title, icon)
        self.mainloop()

    def initVar(self):
        """Initialize variables"""
        self.nameOrder = []
        self.nameDict = {}
        self.smpName = {
            "name": self.nameDict,
            "order": self.nameOrder
        }
        self.initWidgetVar()

    def initWidgetCtrlVar(self):
        """Initialize widget state"""
        self.btnRun["state"] = "disabled"

    def initWindow(self, title, icon):
        """Window initialization"""
        self.title(title)
        try:
            self.iconbitmap(icon)
        except:
            pass
            
        self.resizable(False, False)
        self.createMenu()
        self.content = ttk.Frame(self)
        self.createWidgets()
        self.arrangeWidgets()
        self.initWidgetCtrlVar()
        self.icon = icon

    def createWidgets(self):
        """Create widgets"""
        self.labelFolder = ttk.Label(self.content, text="Assignment Directory:", justify="left")
        self.btnLoadfd = ttk.Button(self.content, text="Select Directory", command=self.loadfolder)
        self.entryFolder = ttk.Entry(self.content, width=Common.ENTRY_WIDTH, state="readonly",
                                     textvariable=self.strFolder)
        
        self.labelxls = ttk.Label(self.content, text="Excel Roster:", justify="left")
        self.btnLoadxls = ttk.Button(self.content, text="Select Roster", command=self.btnLoadxls)
        self.entryXlsfile = ttk.Entry(self.content, width=Common.ENTRY_WIDTH, state="readonly", textvariable=self.strXls)
        self.labelSheetname = ttk.Label(self.content, justify="left", text="Specify Excel sheet name for roster (if empty, use first sheet):")
        self.entrySheetname = ttk.Entry(self.content, textvariable=self.sheetname)
        
        self.createRulesWidgets()
        
        self.labelSample = ttk.Label(self.content, text="Naming Example: (Directories have no extension, spaces in filenames auto-removed)")
        self.entrySample = ttk.Entry(self.content, width=Common.ENTRY_WIDTH * 4 // 5, state="readonly",
                                     foreground="green", textvariable=self.strSample)
        self.updateSample()
        
        self.btnRun = ttk.Button(self.content, text="Start Renaming (Must first specify \"Assignment Directory\" and \"Excel Roster\")", width=Common.ENTRY_WIDTH,
                                 command=self.startRename, state="disabled")

    def createRulesWidgets(self):
        """Create renaming rules widgets"""
        self.lbframeRules = ttk.Labelframe(self.content, text="Renaming Rules (in following order):", 
                                           labelanchor="nw", borderwidth=5, relief="ridge", 
                                           height=200, width=330)
        
        self.labelPrefix1 = ttk.Label(self.lbframeRules, text="1. Prefix1 (leave empty if not needed):", justify="left")
        self.entryPrefix1 = ttk.Entry(self.lbframeRules, validate="focusout", 
                                      validatecommand=self.updateSample, textvariable=self.strPrefix1)
        
        self.cbtnSerialNumber = ttk.Checkbutton(self.lbframeRules, 
                                                text="2. Serial Number (sequential numbering, 3 digits, range: 001~999, e.g.: 017)",
                                                variable=self.bSerno, width=Common.ENTRY_WIDTH - 5,
                                                command=self.updateSample)
        
        self.labelPrefix2 = ttk.Label(self.lbframeRules, text="3. Prefix2 (leave empty if not needed):", justify="left")
        self.entryPrefix2 = ttk.Entry(self.lbframeRules, validate="focusout", 
                                      validatecommand=self.updateSample, textvariable=self.strPrefix2)
        
        self.cbtnStudentNumber = ttk.Checkbutton(self.lbframeRules, text="4. Student ID (e.g., 2025103011)", 
                                                 variable=self.bStuno, width=Common.ENTRY_WIDTH - 5, 
                                                 command=self.updateSample)
        
        self.labelPrefix3 = ttk.Label(self.lbframeRules, text="5. Prefix3 (leave empty if not needed):", justify="left")
        self.entryPrefix3 = ttk.Entry(self.lbframeRules, validate="focusout", 
                                      validatecommand=self.updateSample, textvariable=self.strPrefix3)
        
        self.cbtnName = ttk.Checkbutton(self.lbframeRules, text="6. Name", variable=self.bName,
                                        width=Common.ENTRY_WIDTH - 5, command=self.updateSample, 
                                        state="disabled")
        
        self.labelSuffix = ttk.Label(self.lbframeRules, text="7. Suffix (leave empty if not needed):", justify="left")
        self.entrySuffix = ttk.Entry(self.lbframeRules, validate="focusout", 
                                     validatecommand=self.updateSample, textvariable=self.strSuffix)
        
        self.labelExtention = ttk.Label(self.lbframeRules, 
                                        text="8. Extension (e.g., pdf, leave empty for no change. Directories won't have extension)):", 
                                        justify="left")
        self.entryExtention = ttk.Entry(self.lbframeRules, validate="focusout", 
                                        validatecommand=self.updateSample, textvariable=self.strExtension)
        
        self.labelClass = ttk.Label(self.lbframeRules, 
                                    text="Optional. If has class, insert before position N (1-8) (invalid value means no insertion):", 
                                    justify="left")
        self.entryClass = ttk.Entry(self.lbframeRules, validate="focusout", 
                                    validatecommand=self.updateSample, textvariable=self.strClass)

    def createMenu(self):
        """Create menu"""
        menu = tk.Menu(self)
        self["menu"] = menu
        aboutMenu = tk.Menu(menu, tearoff=0)
        aboutMenu.add_command(label="About", command=self.about)
        aboutMenu.add_command(label="Notice", command=self.notice)
        menu.add_cascade(label="Help", menu=aboutMenu)

    def arrangeWidgets(self):
        """Arrange widgets"""
        self.content.grid(row=0, column=0)
        
        self.labelFolder.grid(row=0, column=0, stick="w")
        self.btnLoadfd.grid(row=0, column=1, stick="e")
        self.entryFolder.grid(row=1, column=0, columnspan=2)
        
        self.labelxls.grid(row=2, column=0, stick="w")
        self.btnLoadxls.grid(row=2, column=1, stick="e")
        self.entryXlsfile.grid(row=3, column=0, columnspan=2)
        
        self.labelSheetname.grid(row=4, column=0, stick="w", padx=(5, 5), pady=(10, 20))
        self.entrySheetname.grid(row=4, column=1, stick="e", padx=(5, 5), pady=(10, 20))
        
        self.arrangeRulesWidgets()
        
        self.labelSample.grid(row=6, column=0, stick="w")
        self.entrySample.grid(row=7, column=0, columnspan=2, stick="n")
        self.btnRun.grid(row=8, column=0, rowspan=2, columnspan=2)

    def arrangeRulesWidgets(self):
        """Arrange renaming rules widgets"""
        self.lbframeRules.grid(row=5, column=0, columnspan=2, pady=(10, 20))
        
        self.labelPrefix1.grid(row=1, column=0, stick="w")
        self.entryPrefix1.grid(row=1, column=1)
        
        self.cbtnSerialNumber.grid(row=2, column=0, columnspan=2)
        
        self.labelPrefix2.grid(row=3, column=0, stick="w")
        self.entryPrefix2.grid(row=3, column=1)
        
        self.cbtnStudentNumber.grid(row=4, column=0, columnspan=2)
        
        self.labelPrefix3.grid(row=5, column=0, stick="w")
        self.entryPrefix3.grid(row=5, column=1)
        
        self.cbtnName.grid(row=6, column=0, columnspan=2)
        
        self.labelSuffix.grid(row=7, column=0, stick="w")
        self.entrySuffix.grid(row=7, column=1)
        
        self.labelExtention.grid(row=8, column=0, stick="w")
        self.entryExtention.grid(row=8, column=1)
        
        self.labelClass.grid(row=9, column=0, stick="w")
        self.entryClass.grid(row=9, column=1)

    def initWidgetVar(self):
        """Initialize widget variables"""
        self.strFolder = tk.StringVar()
        self.strXls = tk.StringVar()
        self.sheetname = tk.StringVar()
        self.strSample = tk.StringVar()
        
        self.strPrefix1 = tk.StringVar()
        self.bSerno = tk.BooleanVar(value=True)
        self.strPrefix2 = tk.StringVar(value='.')
        self.bStuno = tk.BooleanVar(value=True)
        self.strPrefix3 = tk.StringVar(value='.')
        self.bName = tk.BooleanVar(value=True)
        self.strSuffix = tk.StringVar()
        self.strExtension = tk.StringVar()
        self.strClass = tk.StringVar()

    def about(self):
        MsgDialog(self, "about", "关于", self.icon, Common.ABOUT)

    def notice(self):
        MsgDialog(self, "notice", "注意", self.icon, Common.NOTICE, width=self.winfo_width())

    def loadfolder(self):
        self.strFolder.set(filedialog.askdirectory())
        self.updateBtn()

    def btnLoadxls(self):
        self.strXls.set(filedialog.askopenfilename())
        self.updateBtn()

    def updateBtn(self):
        """Update button state"""
        if (len(self.strFolder.get().strip()) != 0) and (len(self.strXls.get().strip()) != 0):
            self.btnRun["state"] = "!disabled"
        else:
            self.btnRun["state"] = "disabled"

    def updateSample(self):
        """Update naming example"""
        self.smpName["name"].clear()
        self.smpName["order"].clear()
        
        prefix1 = self.strPrefix1.get().strip()
        self.smpName["name"]["prefix1"] = prefix1
        self.smpName["order"].append("prefix1")
        
        serialno = True if self.bSerno.get() else ''
        self.smpName["name"]["serialno"] = serialno
        self.smpName["order"].append("serialno")
        
        prefix2 = self.strPrefix2.get().strip()
        self.smpName["name"]["prefix2"] = prefix2
        self.smpName["order"].append("prefix2")
        
        studentno = True if self.bStuno.get() else ''
        self.smpName["name"]["studentno"] = studentno
        self.smpName["order"].append("studentno")
        
        prefix3 = self.strPrefix3.get().strip()
        self.smpName["name"]["prefix3"] = prefix3
        self.smpName["order"].append("prefix3")
        
        studentname = True if self.bName.get() else ''
        self.smpName["name"]["studentname"] = studentname
        self.smpName["order"].append("studentname")
        
        suffix = self.strSuffix.get().strip()
        self.smpName["name"]["suffix"] = suffix
        self.smpName["order"].append("suffix")
        
        extension = self.strExtension.get().strip()
        extension = "." + extension if extension else ""
        self.smpName["name"]["extension"] = extension
        self.smpName["order"].append("extension")
        
        pos = self.strClass.get().strip()
        if pos:
            try:
                pos = int(pos)
                if pos < 1:
                    pos = 1
                elif pos > 8:
                    pos = 8
                self.smpName["name"]["classname"] = True
                self.smpName["order"].insert(pos - 1, "classname")
            except:
                pos = ""
            self.strClass.set(pos)
        else:
            self.smpName["name"]["classname"] = ''
        
        samplename = self.smpName['name'].copy()
        if samplename["serialno"]: samplename["serialno"] = "017"
        if samplename["studentno"]: samplename["studentno"] = "2022030432"
        if samplename["studentname"]: samplename["studentname"] = "Zhang Jia"
        if samplename["classname"]: samplename["classname"] = "22CompSci1"
        if not samplename["extension"]:
            samplename["extension"] = ".smp"
        
        smp = ''
        for v in self.smpName['order']:
            smp += samplename[v]
        self.strSample.set(smp)
        
        return True

    def enableRelatedControls(self, isValid):
        """Enable or disable related controls"""
        if isValid:
            self.entrySheetname['state'] = "!readonly"
            self.entryPrefix1["state"] = "!readonly"
            self.cbtnSerialNumber["state"] = "!disabled"
            self.entryPrefix2["state"] = "!readonly"
            self.cbtnStudentNumber["state"] = "!disabled"
            self.entryPrefix3["state"] = "!readonly"
            self.entrySuffix["state"] = "!readonly"
            self.entryExtention["state"] = "!readonly"
            self.entryClass["state"] = "!readonly"
            self.btnRun["state"] = "!disabled"
        else:
            self.entrySheetname['state'] = "readonly"
            self.entryPrefix1["state"] = "readonly"
            self.cbtnSerialNumber["state"] = "disabled"
            self.entryPrefix2["state"] = "readonly"
            self.cbtnStudentNumber["state"] = "disabled"
            self.entryPrefix3["state"] = "readonly"
            self.entrySuffix["state"] = "readonly"
            self.entryExtention["state"] = "readonly"
            self.entryClass["state"] = "readonly"
            self.btnRun["state"] = "disabled"

    def startRename(self):
        """Start renaming"""
        self.enableRelatedControls(False)
        renameObj = Rename(self.strFolder.get(), self.strXls.get(), self.smpName, self.sheetname.get())
        
        err = renameObj.getErrorInfo()
        if err:
            print('\n' + err + '\n')
        else:
            successinfo = renameObj.getSuccessInfo()
            print('\n' + successinfo + '\n')
        
        self.enableRelatedControls(True)


def main():
    """Main function"""
    title = "Assignment Renaming"
    icon = "1.ico"
    mainwindow = Window(title, icon)


if __name__ == "__main__":
    main()
