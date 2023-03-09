from pptx.shapes.connector import Connector
from pptx.shapes.autoshape import Shape


buffer = 100


class ConnectorDetail:
    connector: Connector
    startAt: Shape
    endAt: Shape

    def __init__(self) -> None:
        self.startAt = None
        self.endAt = None

    def connectorStartsAt(self, shape: Shape) -> bool:
        if self.connector == None:
            return False
        leftTop = (shape.left - buffer, shape.top-buffer)
        rightBottom = (shape.left + shape.width + buffer,
                       shape.top + shape.height + buffer)
        # 左上から右下までの座標内にあれば、OK。実際は、繋がったように見えてるだけのケースもあるから、もうちょい幅が必要
        return (leftTop[0] <= self.connector.begin_x <= rightBottom[0]) & (leftTop[1] <= self.connector.begin_y <= rightBottom[1])

    def connectorEndAt(self, shape: Shape) -> bool:
        if self.connector == None:
            return False
        leftTop = (shape.left - buffer, shape.top - buffer)
        rightBottom = (shape.left + shape.width + buffer,
                       shape.top + shape.height + buffer)
        # 左上から右下までの座標内にあれば、OK。実際は、繋がったように見えてるだけのケースもあるから、もうちょい幅が必要
        return (leftTop[0] <= self.connector.end_x <= rightBottom[0]) & (leftTop[1] <= self.connector.end_y <= rightBottom[1])
