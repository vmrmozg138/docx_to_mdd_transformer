from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

import re

from string import digits
from tkinter import filedialog as tkd

def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within *parent*,
    in document order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
        # print(parent_elm.xml)
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def table_to_array(inputdata):
    resultingarray=[]
    for r in inputdata.rows:
        rdata = [cell.text.strip() for cell in r.cells]
        resultingarray.append(rdata)
    badcols=[]
    cols = inputdata.columns
    if len(cols)>2:
        cnt=0
        for c in cols:
            cdata = [cell.text.strip() for cell in c.cells]
            if len(cdata) != len(set(cdata)):
                badcols.append(cnt)
            cnt+=1
    singlecodepattern = re.compile(r"^[0-9][0-9 ]*$")
    quest = []
    for litem in resultingarray:
        elem = {}
        for x in litem:
            if singlecodepattern.match(x):
                elem.update({'code':x.split(' ')[-1]})
            else:
                if (litem.index(x) not in badcols) and ('label' not in elem) and (len(x)>0):
                    elem.update({'label':x})
        if 'code' in elem and 'label' in elem:
            quest.append(elem)
    return quest
    
def find_1st_cyr_index(text):
    if bool(re.search('[а-яА-Я]', text)):
        cyrstring = 'АаБбВвГгДдЕеЁёЖжЗзИиЙйКкЛлМмНнОоПпРрСсТтУуФфХхЦцЧчШшЩщЪъЫыЬьЭэЮюЯя'
        ind = 0
        for symbol in text:
            ind+=1
            if symbol in cyrstring:
                break
        return ind
    else:
        return -1


file_path_string = tkd.askopenfilename()
docexample = Document(file_path_string)
storage = []

for block in iter_block_items(docexample):
    if isinstance(block, Table):
        storage.append(table_to_array(block))
    if isinstance(block, Paragraph):
        if len(block.text)>0 and find_1st_cyr_index(block.text)>1 and not str(block.text).startswith('/'):
            storage.append(block.text)
    #print(block)
#print(result)

for stitem in storage:
    print('#\n',stitem)
       
with open(file_path_string[:file_path_string.rfind('/')+1]+"metadata.txt", "w") as f:
    for qitem in storage:
        q_output = ''
        if isinstance(qitem, list) and len(qitem)>0:            
            q_output = q_output+"categorial [1..] \n{\n" + ',\n'.join([str("\t_"+i['code'] + " \"" + i['label'] + "\"") for i in qitem]) + "\n};\n\n"
        elif isinstance(qitem, str):
            qname = qitem[:find_1st_cyr_index(qitem)-1].strip()
            print(qname)
            '''
            qlabel = qitem[find_1st_cyr_index(qitem)-1:]
            q_output = q_output + str(qname) + "<b> "+ str(qlabel) +"</b>\n"
            '''
            if re.search(r'[A-Z]+\d', str(qname)):
                qlabel = qitem[find_1st_cyr_index(qitem)-1:].strip()
                qname = qname.replace('.', ' ').strip().replace(' ','_')           
                q_output = q_output + str(qname) + " \"<b>"+ str(qlabel) +"</b>\"\n"
            else:
                q_output = "\n"
        
        f.write(q_output)