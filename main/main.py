import collections.abc
# ↑↑↑のimportの理由は、https://kakakakakku.hatenablog.com/entry/2022/11/24/105451 参照。
# collectionsクラスをPythonが非推奨にした影響の解決法

from pptx import Presentation
from pptx.shapes.autoshape import Shape
from pptx.shapes.connector import Connector

prs = Presentation("test.2.pptx")

for slide in prs.slides:
    for shape in slide.shapes:
        if isinstance(shape, Shape):
            print("SHAPE: id&text: " + str(shape.shape_id) + " " + shape.text)
            print("top: " + str(shape.top))
            print("left: " + str(shape.left))
            print("\n")
            continue

        elif isinstance(shape, Connector):
            print("CONNECTOR: shape_id: " + str(shape.shape_id))
            print("begin: (" + str(shape.begin_x) +
                  "," + str(shape.begin_y) + ")")
            print("end: (" + str(shape.end_x) +
                  "," + str(shape.end_y) + ")")
            print("height: " + str(shape.height))
            print("width: " + str(shape.width))
            print("\n")
            continue
        else:
            print("other shape detected!")
            print(shape)
            print("\n")

print("finish!")
