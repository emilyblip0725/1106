from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from pathlib import Path

OUT = Path(r'C:\Users\ncucc\Desktop\1030\schedule-task-prototype\個人行程與任務管理系統_PRD.docx')
doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.72); sec.bottom_margin = Inches(.7); sec.left_margin = Inches(.78); sec.right_margin = Inches(.78)
sec.page_width = Inches(8.5); sec.page_height = Inches(11)

# Theme and styles
styles = doc.styles
for name in ['Normal','Title','Subtitle','Heading 1','Heading 2','Heading 3']:
    st=styles[name]
    st.font.name='Aptos'
    st._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
    st.font.color.rgb=RGBColor(0,0,0)
styles['Normal'].font.size=Pt(10.5)
styles['Normal'].paragraph_format.space_after=Pt(5)
styles['Normal'].paragraph_format.line_spacing=1.13
styles['Title'].font.size=Pt(25); styles['Title'].font.bold=True
styles['Title'].paragraph_format.space_after=Pt(8)
styles['Subtitle'].font.size=Pt(11); styles['Subtitle'].font.color.rgb=RGBColor(92,105,98)
styles['Heading 1'].font.size=Pt(16); styles['Heading 1'].font.bold=True
styles['Heading 1'].paragraph_format.space_before=Pt(15); styles['Heading 1'].paragraph_format.space_after=Pt(6)
styles['Heading 2'].font.size=Pt(12); styles['Heading 2'].font.bold=True
styles['Heading 2'].paragraph_format.space_before=Pt(9); styles['Heading 2'].paragraph_format.space_after=Pt(4)
styles['Heading 3'].font.size=Pt(10.5); styles['Heading 3'].font.bold=True

# Table helper
def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)
def cell_margins(cell, top=95, start=110, bottom=95, end=110):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr(); mar=OxmlElement('w:tcMar')
    for m,v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        node=OxmlElement('w:'+m);node.set(qn('w:w'),str(v));node.set(qn('w:type'),'dxa');mar.append(node)
    tcPr.append(mar)
def set_borders(table):
    tblPr=table._tbl.tblPr; borders=OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        el=OxmlElement('w:'+edge);el.set(qn('w:val'),'single');el.set(qn('w:sz'),'5');el.set(qn('w:space'),'0');el.set(qn('w:color'),'D9E0DC');borders.append(el)
    tblPr.append(borders)
def table(headers, rows, widths=None, font=9.2):
    t=doc.add_table(rows=1, cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    if widths:
        for c,w in zip(t.columns,widths): c.width=Inches(w)
    for i,h in enumerate(headers):
        cell=t.rows[0].cells[i];cell.text=h;shade(cell,'3E6252');cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell_margins(cell)
        for p in cell.paragraphs:
            p.paragraph_format.space_after=Pt(0)
            for r in p.runs:r.font.bold=True;r.font.size=Pt(font);r.font.color.rgb=RGBColor(255,255,255);r.font.name='Aptos';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        if widths:
            for c,w in zip(cells,widths):c.width=Inches(w)
        for i,val in enumerate(row):
            cells[i].text=str(val);cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;cell_margins(cells[i])
            if ri%2==1:shade(cells[i],'F3F7F4')
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.08
                for r in p.runs:r.font.size=Pt(font);r.font.name='Aptos';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
    set_borders(t);doc.add_paragraph().paragraph_format.space_after=Pt(1)
    return t
def para(text='', boldlead=None):
    p=doc.add_paragraph()
    if boldlead and text.startswith(boldlead):
        p.add_run(boldlead).bold=True;p.add_run(text[len(boldlead):])
    else:p.add_run(text)
    return p
def bullets(items):
    for item in items:
        p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(item)
def numbered(items):
    for item in items:
        p=doc.add_paragraph(style='List Number');p.paragraph_format.space_after=Pt(2);p.add_run(item)

# Header and footer
header=sec.header.paragraphs[0];header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
rr=header.add_run('產品需求文件  /  拾序');rr.font.size=Pt(8);rr.font.color.rgb=RGBColor(115,132,122)
footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=footer.add_run('個人行程與任務管理系統  ·  PRD  ·  ');r.font.size=Pt(8);r.font.color.rgb=RGBColor(130,142,135)
fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');footer._p.append(fld)

p=doc.add_paragraph(style='Title');p.add_run('個人行程與任務管理系統')
p=doc.add_paragraph(style='Subtitle');p.add_run('產品需求文件  Product Requirements Document')
p=doc.add_paragraph();p.paragraph_format.space_before=Pt(10);p.paragraph_format.space_after=Pt(12)
r=p.add_run('文件版本  1.0    |    狀態  初版需求基線    |    日期  2026 年 10 月 30 日');r.font.size=Pt(9);r.font.color.rgb=RGBColor(100,115,106)
para('本文件定義一套以電腦使用為優先的個人行程與任務管理產品，核心目標是將分散的行程與待辦集中呈現，協助使用者辨識先後順序並降低遺漏。需求依目前訪談與互動原型整理；尚未確認的部分明確列為假設或待決事項，供後續驗證。')

doc.add_heading('一 產品概述',level=1)
para('產品暫名「拾序」，提供個人使用者一個集中管理每日行程與待辦事項的工作空間。首頁以今日總覽呈現任務、行程、優先事項與完成進度，並延伸提供所有任務清單、月曆、搜尋與專注計時。')
table(['欄位','說明'],[
('產品型態','個人行程與任務管理系統；第一階段以桌面瀏覽器前端為主。'),
('主要使用者','希望將生活或工作行程、待辦集中管理的個人使用者；實際職業情境尚未確認。'),
('主要痛點','行程與任務分散；容易忘記待辦；難以安排優先順序。'),
('產品原則','今天先看得懂、重要事項更醒目、快速記錄、逐步完成、不以高壓提醒干擾使用者。'),
('目前原型範圍','桌面優先的互動式前端；支援新增、編輯、完成任務與管理行程，資料暫存於瀏覽器本機。')
],[1.5,5.75])

doc.add_heading('二 目標與成功指標',level=1)
doc.add_heading('產品目標',level=2)
bullets(['集中查看同一天的行程與待辦，減少在不同工具之間切換。','以優先級、到期日和每日總覽協助使用者決定下一步。','讓新增、修改、完成與查找事項的操作足夠直接。','以溫和、可預期的方式呈現提醒，不在未取得同意前發送系統通知。'])
doc.add_heading('建議驗收／成效指標',level=2)
table(['指標','定義與建議目標','驗證方式'],[
('快速記錄成功率','測試者可在 30 秒內新增一筆任務或行程；目標 ≥ 90%。','可用性測試任務觀察。'),
('今日總覽理解度','測試者可在 10 秒內指出今日優先事項與下一個行程；目標 ≥ 85%。','原型走查與理解題。'),
('核心流程完成率','新增、完成、編輯及搜尋任務的流程完成率；目標 ≥ 90%。','可用性測試事件記錄。'),
('資料留存可靠性','重新整理頁面後，本機新增資料仍可讀取；目標 100%。','功能驗收測試。'),
('遺漏自評改善','使用者使用一至兩週後，自評忘記待辦的頻率下降。','後續訪談或簡短問卷；尚待建立基線。')
],[1.35,3.1,2.8])

doc.add_heading('三 使用者與使用情境',level=1)
para('目前確認使用裝置為電腦。使用者希望處理工作與個人事項交錯、容易漏看待辦，以及不易判斷優先次序的問題。職業、是否需要跨裝置同步、提醒時機與實際日程型態仍待確認。')
doc.add_heading('核心情境',level=2)
numbered(['使用者開啟系統，快速確認今天的行程、待辦、優先事項與完成狀態。','使用者新增一筆任務，指定類別、到期日與優先級；之後在清單中標記完成或修改內容。','使用者建立行程並指定日期、時間、結束時間及地點，在今日時間軸或月曆中查看。','使用者搜尋既有任務或行程，快速定位並開啟編輯。','使用者開始一段專注時段，暫時專注在單一工作項目。'])

doc.add_heading('四 產品範圍',level=1)
table(['優先級','功能範圍','說明'],[
('P0 核心','今日總覽','同頁呈現當日摘要、未完成優先任務、今日行程與每週節奏。'),
('P0 核心','任務管理','新增、編輯、完成／恢復；具名稱、備註、類別、到期日、優先級。'),
('P0 核心','行程管理','新增、編輯、刪除行程；具名稱、日期、開始／結束時間、地點或備註。'),
('P0 核心','任務清單與篩選','查看所有任務，依工作、生活、學習類別篩選，依日期排序。'),
('P0 核心','行事曆','月視圖查看行程及未完成任務；選日期可查看當日總覽。'),
('P1 輔助','搜尋','以文字搜尋任務名稱、備註、類別及行程名稱。'),
('P1 輔助','提醒展示','原型提供提醒面板以驗證內容與位置；不代表已有推播或排程通知能力。'),
('P1 輔助','專注計時','提供 25 分鐘專注計時器，可開始、暫停、繼續與重設。'),
('P2 待評估','多裝置同步、帳號、雲端備份、重複任務、通知排程、外部日曆整合','需在技術方案與使用者需求確認後另行定義。')
],[1.15,2.05,4.05])
para('不在目前原型範圍：多人協作、分享行事曆、電子郵件／訊息整合、AI 自動排程、原生手機應用，以及真實瀏覽器或作業系統通知。')

doc.add_heading('五 功能需求',level=1)
doc.add_heading('FR-01 今日總覽',level=2)
bullets(['首頁顯示選定日期的行程數、任務數、未完成高優先事項數與任務完成率。','首頁列出當日未完成任務，優先級排序為高、一般、有空再做；同級以時間排序。','首頁以時間順序顯示當日行程，呈現時間、名稱及地點或備註。','可由首頁快速新增任務或行程，並可跳轉至任務清單或行事曆。','日期控制項變更後，摘要、任務與行程內容須同步更新。'])
doc.add_heading('FR-02 任務管理',level=2)
bullets(['任務欄位：名稱（必填，最多 80 字）、備註（選填，最多 200 字）、類別、到期日、時間、優先級與完成狀態。','預設類別為「工作」「生活」「學習」；系統需支援後續新增類別。','優先級至少包含「高優先」「一般」「有空再做」，建立任務預設為一般。','任務可新增、編輯、標記完成及恢復為未完成；完成後需有明確視覺狀態。','所有任務頁提供類別篩選與日期排序；空清單時提供清楚的空狀態與新增入口。'])
doc.add_heading('FR-03 行程管理',level=2)
bullets(['行程欄位：名稱（必填）、日期（必填）、開始時間、結束時間、地點與備註。','行程可新增、編輯與刪除；刪除需有可辨識的操作入口，正式產品建議增加確認或復原機制。','今日總覽依開始時間排序行程；月曆顯示行程名稱及時間摘要。','應驗證結束時間不早於開始時間；原型目前需補強此項資料驗證。'])
doc.add_heading('FR-04 搜尋與導航',level=2)
bullets(['提供任務及行程的共用搜尋入口；輸入關鍵字後即時顯示最多九筆匹配結果。','搜尋結果須標明項目類型與類別或日期，選取後開啟對應編輯表單。','支援主要頁面切換：今日總覽、所有任務、行事曆。','鍵盤快捷鍵：Ctrl／Command + K 開啟搜尋；N 開啟新增任務；Esc 關閉浮層。'])
doc.add_heading('FR-05 提醒與專注',level=2)
bullets(['原型提醒面板提供今日到期任務與近期行程示例，並清楚標註此為介面展示。','本機原型不發送電子郵件、瀏覽器推播或作業系統通知。','專注計時預設 25 分鐘，提供開始、暫停／繼續、重設與完成提示。'])

doc.add_heading('六 主要使用流程',level=1)
table(['流程','步驟','完成條件'],[
('新增任務','首頁按「新增任務」→填名稱／日期／類別／優先級→儲存。','任務出現在對應日期摘要與所有任務清單。'),
('完成任務','在任務列勾選完成。','完成狀態立即更新，摘要進度同步重算。'),
('新增行程','首頁或行事曆按「新增行程」→填名稱／日期／時間等→儲存。','行程出現在當日時間軸與對應月曆日期。'),
('搜尋項目','按搜尋或 Ctrl／Command + K→輸入關鍵字→選取結果。','開啟該項目的編輯表單。'),
('查看日期','使用日期選擇器、週圖表或月曆選日期。','今日總覽顯示該日期的任務與行程。')
],[1.2,3.7,2.4])

doc.add_heading('七 資料需求與規則',level=1)
table(['資料實體','欄位','規則'],[
('任務 Task','id、title、note、date、time、category、priority、done','title、date 必填；priority ∈ high／medium／low；done 為布林值。'),
('行程 Event','id、title、note、date、time、endTime、location、tone','title、date 必填；開始與結束時間需符合時間順序；tone 為展示樣式屬性。'),
('類別 Category','name、color','預設工作、生活、學習；類別名稱須唯一。')
],[1.15,2.75,3.4])
para('原型目前將資料存於瀏覽器 localStorage。清除瀏覽器資料或改用其他瀏覽器／裝置時，資料不會自動同步或復原。正式產品需評估登入、雲端儲存、備份匯出及資料刪除策略。')

doc.add_heading('八 非功能需求',level=1)
bullets(['易用性：主要新增入口在今日總覽及對應清單可見；表單欄位標籤清楚並可用鍵盤操作。','回應性：一般瀏覽器上主要操作應即時回應；搜尋輸入後應快速呈現結果。','相容性：桌面優先，版面需可在窄螢幕維持可讀；正式支援瀏覽器版本待技術決策。','可存取性：提供表單標籤、按鈕名稱、焦點狀態及足夠色彩對比；完成狀態不得只以顏色表達。','隱私：任務與行程可能含個人資訊；正式產品需說明資料儲存位置、保留及刪除方式。','可靠性：本機原型刷新後保留資料；正式產品需定義備份、同步衝突及故障復原策略。'])

doc.add_heading('九 驗收條件',level=1)
numbered(['使用者可在今日總覽看到選定日期的行程、未完成任務及摘要數字。','新增任務後，任務清單及相同日期的首頁摘要立即反映新資料。','任務完成或恢復後，完成狀態、待辦數與完成率正確更新。','新增或編輯行程後，當日時間軸及月曆內容正確更新。','在所有任務頁選擇工作、生活或學習，清單只呈現對應類別。','搜尋任務或行程名稱可找到結果；選取後開啟正確項目編輯表單。','瀏覽器重新整理後，已儲存的本機資料仍可讀取。','提醒面板明確說明為原型展示，不暗示已送出系統通知。','桌面及窄螢幕版面不出現文字重疊或主要操作被裁切。'])

doc.add_heading('十 假設、限制與待決事項',level=1)
table(['項目','目前假設／限制','需確認的問題'],[
('使用者族群','目前按個人使用者設計，工作、生活、學習為預設分類。','主要使用者是上班族、學生、自由工作者，還是其他族群？'),
('裝置','已確認以電腦使用為主；原型雖有窄版響應式，尚未確認手機為正式需求。','是否需要手機使用及跨裝置一致資料？'),
('提醒','到期日、優先級與介面內今日提醒作為第一版方向；無真實推播。','希望採用桌面通知、郵件、每日摘要或僅系統內提醒？提醒提前多久？'),
('資料儲存','原型只使用瀏覽器本機儲存。','正式產品是否需要帳號、雲端同步、匯出與備份？'),
('日曆範圍','原型提供月曆及日檢視切換；未整合外部行事曆。','是否需週／日視圖、重複行程或 Google／Outlook 同步？'),
('時間格式','以瀏覽器本地日期時間顯示，尚未定義時區與全天事件規則。','是否要支援全天行程、跨時區及自訂週起始日？'),
('產品優先順序','目前以原型驗證核心工作流程，未定義正式上線日期或技術架構。','下一階段優先精修哪些流程與視覺？')
],[1.15,3.0,3.15])

doc.add_heading('十一 版本規劃建議',level=1)
table(['階段','建議內容','退出條件'],[
('原型驗證','以此互動原型測試今日總覽、任務排序、行程新增、搜尋及專注流程。','完成至少 5 位目標使用者測試；確認資訊架構及核心表單。'),
('MVP','正式任務／行程資料模型、穩定持久化、基本提醒設定、無障礙與跨瀏覽器驗收。','核心驗收條件通過；資料保存與提醒行為經使用者確認。'),
('後續迭代','跨裝置同步、外部日曆、重複事項、進階通知、資料匯出與分析。','依使用數據和需求訪談排序，不預設全數納入。')
],[1.2,3.55,2.55])

# Core properties
props=doc.core_properties;props.title='個人行程與任務管理系統 產品需求文件';props.subject='PRD';props.author='產品規劃';props.keywords='PRD,行程管理,任務管理'
doc.save(OUT)
print(OUT)
