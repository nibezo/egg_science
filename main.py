from openpyxl import load_workbook

wb = load_workbook('EEG Science.xlsx')
eeg = wb.active
title, paperAbstract, fieldsOfStudy, pdfUrls, authors, venue, year_published, inCitations = [], [], [], [], [], [], [], []
f = open('egg_big.md', 'a+', encoding="utf-8")
f.write("|title|paperAbstract|fieldsOfStudy|pdfUrls|authors|venue|year_published|inCitations|\n")
f.write("|:----:|:----:|:----:|:----:|:----:|:----:|:----:|:----:|\n")

for _ in range(1, 75867):
    title.append(eeg[f'A{_}'].value)
    paperAbstract.append(eeg[f'B{_}'].value)
    fieldsOfStudy.append(eeg[f'C{_}'].value)
    pdfUrls.append(eeg[f'D{_}'].value)
    authors.append(eeg[f'F{_}'].value)
    venue.append(eeg[f'E{_}'].value)
    year_published.append(eeg[f'G{_}'].value)
    inCitations.append(eeg[f'H{_}'].value)
    print(_)

for _ in range(0, 75866):
    if len(str(paperAbstract[_])) >= 180:
        paperAbstract[_] = paperAbstract[_][0:150] + '...'

    paperAbstract[_] = str(paperAbstract[_]).replace('|', ' ')

    title[_] = str(title[_]).replace('|', '')
    fieldsOfStudy[_] = str(fieldsOfStudy[_]).replace('|', ' ')
    authors[_] = str(authors[_]).replace('|', '')
    venue[_] = str(venue[_]).replace('|', '')
    paperAbstract[_] = str(paperAbstract[_]).replace('\n', ' ')
    title[_] = str(title[_]).replace('\n', '')
    fieldsOfStudy[_] = str(fieldsOfStudy[_]).replace('\n', ' ')
    authors[_] = str(authors[_]).replace('\n', '')
    venue[_] = str(venue[_]).replace('\n', '')
    print(_)
for _ in range(1, 75866):
    f.write(f"|{title[_]}|{paperAbstract[_]}|{fieldsOfStudy[_]}|{pdfUrls[_]}|{authors[_]}|{venue[_]}|{year_published[_]}|{inCitations[_]}|\n")
    print(_)

f.close()
