import openpyxl
from qgis.core import QgsVectorLayer, QgsFeature, QgsPointXY, QgsGeometry, QgsProject

# Excelファイルを読み取ります
excel_file_path = r'Excelファイルのパス' //中身を自分のパスに書き換えてね
wb = openpyxl.load_workbook(excel_file_path)
sheet = wb.active

# QGISプロジェクトをロードします
project = QgsProject.instance()
project.read(r'QGISのプロジェクトファイルのパス')//中身を自分のパスに書き換えてね

# マップレイヤーを取得します
layer = project.mapLayersByName('追加したレイヤー名')[0]  # レイヤー名を適切に変更してください

# Excelからデータを読み取り、プロットします
for row in sheet.iter_rows(min_row=2, values_only=True):
x, y = row[1], row[2]  # Excelの列に応じて調整
point = QgsPointXY(x, y)

# QgsPointXY オブジェクトを QgsGeometry オブジェクトに変換
geometry = QgsGeometry.fromPointXY(point)

feature = QgsFeature()
feature.setGeometry(geometry)
layer.dataProvider().addFeatures([feature])

# マップを更新します
layer.triggerRepaint()

# プロジェクトを保存します
project.write()

print('データの読み取り、プロット、アップロードが完了しました。')
