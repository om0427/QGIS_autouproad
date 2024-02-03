import openpyxl
from qgis.core import QgsVectorLayer, QgsFeature, QgsPointXY, QgsGeometry, QgsProject

# Excelファイルを読み取ります
excel_file_path = r'C:\仕事\課題\R5_I4\地域学\outoveexcel.xlsx'
wb = openpyxl.load_workbook(excel_file_path)
sheet = wb.active

# QGISプロジェクトをロードします
project = QgsProject.instance()
project.read(r'C:\仕事\課題\R5_I4\地域学\outo_v2.qgs')

# マップレイヤーを取得します
layer = project.mapLayersByName('outoveexcel-data1')[0]  # レイヤー名を適切に変更してください

# Excelからデータを読み取り、プロットします
for row in sheet.iter_rows(min_row=2, values_only=True):
    x, y = row[1], row[2]  # Excelの列に応じて調整
    
    # xとyがNoneでないことを確認
    if x is not None and y is not None:
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
