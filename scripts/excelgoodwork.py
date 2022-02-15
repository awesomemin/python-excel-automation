import openpyxl

sourceFileAddress = r'.\files\2021\사료 검정기록서(21.03).xlsx'
revisedFileAddress = r'.\files\2021 검정대장\◆검정대장 3개월치 빈것.xlsx'

wb1 = openpyxl.load_workbook(sourceFileAddress)
wb2 = openpyxl.load_workbook(revisedFileAddress)

ws1 = wb1['Sheet1 (2)']
ws2 = wb2['Sheet1']

validrows = []
examineitemnumbers = []

numToIngredient = {1:'A', 2:'B', 3:'수분', 4:'조회분', 5:'조단백', 6:'조지방', 7:'조섬유', 8:'NDF', 9:'ADF', 10:'펩신소화율', 11:'염분', 12:'불소', 13:'칼슘', 14:'인', 15:'납', 16:'크롬', 17:'비소', 18:'수은', 19:'카드뮴', 20:'주석'}
numToCol = {1:'A', 2:'B', 3:'C', 4:'D', 5:'E', 6:'F', 7:'G', 8:'H', 9:'I', 10:'J', 11:'K', 12:'L', 13:'M', 14:'N', 15:'O', 16:'P', 17:'Q', 18:'R', 19:'S', 20:'T'}

page=1
while(page <= 120):
  col = ws1['A']
  for cell in col:
    if(cell.row >= 10+(29*(page-1)) and cell.row <= 28+(29*(page-1))):
      if(cell.value != None):
        validrows.append(cell)
  page = page + 1

for validrow in validrows:
  row = ws1[validrow.row]
  examineitemnumber = -2
  for cell in row:
    if(cell.value != None):
      examineitemnumber = examineitemnumber + 1
  examineitemnumbers.append(examineitemnumber)

index = 0
page = 1
remainingLines = 29
for item in validrows:
  requiredLines = examineitemnumbers[index]
  #어떤 항목 검사했는지 알 수 있는 부분
  itemrow = ws1[item.row]
  times = 0
  examinedItems = []
  for cell in itemrow:
    times = times + 1
    if (times <= 2):
      continue
    else:
      if(cell.value != None):
        examinedItems.append(cell.column)

  if (requiredLines > remainingLines):
    page = page + 1
    remainingLines = 29
  ws2[f'A{7+(29-remainingLines)+(36 * (page-1))}'] = item.value
  #성분명, 검정결과 삽입하는 부분
  index2 = 0
  for plz in examinedItems:
    if index2 == 0:
      ws2[f'C{7+(29-remainingLines)+(36 * (page - 1)) + index2}'] = ws1[f'B{item.row}'].value
    ws2[f'D{7+(29-remainingLines)+(36 * (page - 1)) + index2}'] = numToIngredient[plz]
    ws2[f'H{7+(29-remainingLines)+(36 * (page - 1)) + index2}'] = numToIngredient[plz]
    ws2[f'I{7+(29-remainingLines)+(36 * (page - 1)) + index2}'] = ws1[f'{numToCol[plz]}{item.row}'].value
    index2 = index2 + 1
  remainingLines = remainingLines - requiredLines
  if (requiredLines > remainingLines):
    page = page + 1
    remainingLines = 29
  index = index + 1


wb2.save(revisedFileAddress)

#print(number)