import errno, os, re, sys, tkinter as tk
import configparser as cp, cv2, numpy as np, openpyxl, pandas as pd, pytesseract
from openpyxl.styles import Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import ttk
from tkinter.filedialog import askopenfilename, askopenfilenames, asksaveasfilename

class TextRedirector(object):
    def __init__(self, tb, tag= 'stdout'):
        self.tb = tb
        self.tag = tag

    def write(self, str):
        self.tb.configure(state= 'normal')
        self.tb.insert('end', str, (self.tag,))
        self.tb.update_idletasks()
        self.tb.see(tk.END)
        self.tb.configure(state= 'disabled')

def generateDropLog(tb, xltf, dlf, kwf, isVerbose):
    bosses = ['Lotus', 'Damien', 'Lucid', 'Will', 'Divine King Slime', 'Dusk', 'Djunkel', 'Heretic Hilla', 'Black Mage', 'Seren', 'Kalos', 'Kaling']  
    xltfile = xltf.get()
    dlfile = dlf.get()
    kwfile = kwf.get()
    keywords = []
    images = []
    
    tb.configure(state="normal")
    tb.delete(1.0, tk.END)
    tb.configure(state="disabled")
    
    wb, ws, df = createDataFrame(xltfile)
    
    try:
        images.extend(askopenfilenames(title='Select image files to process', initialdir=os.getcwd(), filetypes=[('Image files', '.jpg .jpeg .png .bmp .tiff')]))
        if len(images) == 0:
            raise FileNotFoundError('No image file(s) selected!')
    except FileNotFoundError as e:
        print(e)
        return
    with open(kwfile := 'keywords.txt') as kw:
        for line in kw:
            keywords.append(line.rstrip('\n'))
    
    for img_path in images:
        drop_list = {}
        print('Reading in', img_name := os.path.basename(img_path), '... ', end='')
        img = cv2.imread(img_path)
        print('Done!')
        
        # Image pre-processing: Greyscaling, Otsu's thresholding, resizing, Otsu's thresholding again, Gaussian blurring
        print('Processing', img_name, '... ', end='')
        inv = np.invert(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
        ot, ot_result = (otsu := lambda input : cv2.threshold(input, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU))(inv)
        scale_factor = 4
        width = int(img.shape[1] * scale_factor)
        height = int(img.shape[0] * scale_factor)
        rsz = cv2.resize(ot_result, (width, height), interpolation = cv2.INTER_AREA)
        ot_2, ot_result_2 = otsu(rsz)
        blur = cv2.GaussianBlur(ot_result_2,(5,5),0)
        print('Done!')
        
        # Pass processed image to Tesseract OCR
        print('Detecting text in', img_name, '... ', end='')
        output = pytesseract.image_to_string(blur, lang='eng',config='--psm 4 --oem 1')
        rep = re.sub('v¥|¥V|WY|YY|VV|VY|vY|WV|vv|\"W|\'\'W', 'W', output) # This regex corrects the problematic capital W in the source image
        drops = rep.split('\n')
        drops = list(filter(None, drops))
        print('Done!')
        
        if isVerbose is True:
            print('Tesseract output:', dash := '----------------------------------------', *drops, dash, sep='\n')
        
        with open(dlfile := 'droplist.txt') as dl:
            for line in dl:
                drop_list[line.rstrip('\n')] = 0
        
        for b in bosses:
            if b.casefold() in img_name.casefold():
                print('Recording drops from', b, '... ', end='')
                
                # Retrieve drop quantities from Tesseract output
                for kw in keywords:
                    for dr in drops:
                        if f'{kw}' in dr:
                            qnt = re.search('[x]{1}\d+', dr)
                            quantity = int(qnt.group(0).split('x')[1])
                            for key, val in drop_list.items():
                                if f'{kw}' in key: drop_list[key] = val + quantity
                
                # Record drop quantities in DataFrame
                for row in range(len(drop_list)):
                    for key, val in drop_list.items():
                        if f'{key}' in df.at[row + 1, 'Item']:
                            df.at[row + 1, b] = val + df.at[row + 1, b] if (df.at[row + 1, b] is not None) else val
                print('Done!\n')
    
    if isVerbose is True:
        with pd.option_context('expand_frame_repr', False, 'display.max_rows', None, 'display.max_columns', None, 'display.width', 0, 'display.max_colwidth', 0):
            print(df, end= '\n\n')
    
    # Write DataFrame contents to Excel worksheet
    rows = dataframe_to_rows(df, index= False)
    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row, 1):
            ws.cell(row= r, column= c, value= value).alignment = Alignment(horizontal= 'center')
    wb.template = False
    print('Writing drop data to', (os.path.basename(xl := asksaveasfilename(title= 'Save Excel file', filetypes= [('Excel Workbook', '.xlsx'), ('Excel 97- Excel 2003 Workbook', '.xls')], defaultextension= '.xlsx'))),'... ', end='')
    wb.save(xl)
    print('Done!')
  
# Create a Pandas DataFrame using an Excel template  
def createDataFrame(xltfile):
    wb = openpyxl.load_workbook(xltfile)
    print('Excel template loaded:', os.path.basename(xltfile), '\n')
    ws = wb[wb.sheetnames[0]]
    df = pd.DataFrame(ws.values)
    columnNames = df.iloc[0]
    df = df[1:]
    df.columns = columnNames
    return (wb, ws, df)

def readFile(cfg, label, filevar, filename, Title, types):
    filename.set(askopenfilename(title= Title, initialdir= os.getcwd(), filetypes= types))
    cfg.set('Files', filevar, filename.get())
    label.config(text= filename.get())
    with open('dlconfig.ini', 'w') as cfgfile:
        cfg.write(cfgfile)

def main():
    window = tk.Tk()
    window.title('Drop Logging Tool')
    window.minsize(971, 600)
    window.resizable(True, True)
    frame1 = tk.Frame(master= window)
    frame2 = tk.Frame(master= window)
    frame3 = tk.Frame(master= window)
    isVerbose = tk.BooleanVar(window)
    xltfile = tk.StringVar(window)
    dlfile = tk.StringVar(window)
    kwfile = tk.StringVar(window)
    cfg = cp.ConfigParser()
    
    isVerbose.set(False)
    
    try:
        cfg.read('dlconfig.ini')
        files = cfg['Files']
        xltfile.set(files['xltfile'])
        dlfile.set(files['dlfile'])
        kwfile.set(files['kwfile'])
    except (FileNotFoundError, KeyError) as e:
        cfg.add_section('Files')
        cfg.set('Files', 'xltfile', '')
        cfg.set('Files', 'dlfile', '')
        cfg.set('Files', 'kwfile', '')
        with open('dlconfig.ini', 'w') as cfgfile:
            cfg.write(cfgfile)
    
    showCurrentExcelLabel = ttk.Label(master= frame2, text= xltfile.get())
    showCurrentDropListLabel = ttk.Label(master= frame2, text= dlfile.get())
    showCurrentKeywordsLabel = ttk.Label(master= frame2, text= kwfile.get())
    
    loadExcelButton = ttk.Button(master= frame1, text= 'Load Excel template', width= 30, command= lambda: readFile(cfg, showCurrentExcelLabel, 'xltfile', xltfile, 'Select Excel template', [('Template', '.xltx'), ('Template (code)', '.xltm'), ('Excel Workbook', '.xlsx'), ('Excel 97- Excel 2003 Workbook', '.xls')]))
    loadDropListButton = ttk.Button(master= frame1, text= 'Load drop list', width= 30, command= lambda: readFile(cfg, showCurrentDropListLabel, 'dlfile', dlfile, 'Select drop list text file', [('Text file', '.txt'), ('All files', '.*')]))
    loadKeywordsButton = ttk.Button(master= frame1, text= 'Load keywords', width= 30, command= lambda: readFile(cfg, showCurrentKeywordsLabel, 'kwfile', kwfile, 'Select keywords text file', [('Text file', '.txt'), ('All files', '.*')]))
    isVerboseButton = ttk.Checkbutton(master= frame1, text= 'Verbose', variable= isVerbose)
    
    generateDropLogButton = ttk.Button(master= frame1, text= 'Generate Excel sheet', width= 30, command= lambda: generateDropLog(textbox, xltfile, dlfile, kwfile, isVerbose.get()))
    textbox = tk.Text(master= frame3, font= ('Consolas', 10), wrap= tk.NONE, state= 'disabled')
    vert_scrollbar = ttk.Scrollbar(master= frame3, orient= 'vertical', command= textbox.yview)
    hori_scrollbar = ttk.Scrollbar(master= frame3, orient= 'horizontal', command= textbox.xview)
    
    frame1.pack()
    frame2.pack()
    frame3.pack(side= 'bottom', fill= 'both', expand= True)
    
    loadExcelButton.grid(row= 1, column= 1)
    loadDropListButton.grid(row= 1, column= 2)
    loadKeywordsButton.grid(row= 1, column= 3)
    generateDropLogButton.grid(row= 2, column= 1, columnspan= 2)
    isVerboseButton.grid(row= 2, column= 3, sticky= 'w')
    
    frame2.columnconfigure(1, weight= 1)
    frame2.columnconfigure(2, weight= 1)
    ttk.Label(master= frame2, text= 'Selected Excel template:', width= 25).grid(row= 1, column= 1, sticky= 'w')
    showCurrentExcelLabel.grid(row= 1, column= 2, sticky= 'w')
    ttk.Label(master= frame2, text= 'Selected drop list file:', width= 25).grid(row= 2, column= 1, sticky= 'w')
    showCurrentDropListLabel.grid(row= 2, column= 2, sticky= 'w')
    ttk.Label(master= frame2, text= 'Selected keywords file:', width= 25).grid(row= 3, column= 1, sticky= 'w')
    showCurrentKeywordsLabel.grid(row= 3, column= 2, sticky= 'w')
    
    frame3.rowconfigure(0, weight= 1)
    frame3.columnconfigure(0, weight= 1)
    textbox.grid(row= 0, column= 0, sticky= 'nsew')
    vert_scrollbar.grid(row= 0, column= 1, sticky= 'ns')
    textbox['yscrollcommand'] = vert_scrollbar.set
    hori_scrollbar.grid(row= 1, column= 0, sticky= 'ew')
    textbox['xscrollcommand'] = hori_scrollbar.set
    sys.stdout = TextRedirector(textbox, 'stdout')
    sys.stderr = TextRedirector(textbox, 'stderr')
    
    print('Ready')
    window.mainloop()

if __name__ == "__main__":
    main()
