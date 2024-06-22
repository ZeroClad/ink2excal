from xml.etree import ElementTree
import json
import win32clipboard as wcb
import random


class InkParser:
    def __init__(self, data):
        # edit config here
        self.overall_scale = 0.04
        self.force_ignore_pressure = False
        self.pen_scale = 6
        self.sample_rate = 1
        self.export_to_file = False
        self.file_path = "./a.json"
        # config end

        self.ns = {"ink": "http://www.w3.org/2003/InkML"}
        self.brushes = {}
        self.data = data
        self.traces: list[Trace] = []
        self.json = ""

    def read_inkml(self):
        tree = ElementTree.fromstring(self.data)
        for context in tree.findall('*/ink:brush', self.ns):
            b = Brush()
            for property in context.findall('ink:brushProperty', self.ns):
                if property.get('name') == "width":
                    b.width = float(property.get('value'))
                elif property.get('name') == "height":
                    b.height = float(property.get('value'))
                elif property.get('name') == "color":
                    b.color = property.get('value')
                elif property.get('name') == "transparency":
                    b.transparency = int(property.get('value'))
                elif property.get('name') == "ignorePressure":
                    b.ignore_pressure = int(property.get('value'))
            self.brushes[context.get('{http://www.w3.org/XML/1998/namespace}id')] = b

        for line in tree.findall('ink:trace', self.ns):
            trace = Trace()
            trace.brush = self.brushes[line.get('brushRef').replace("#", "")]
            ds = line.text.split(',')
            state = 0
            c = Node()      # coordinate
            v = Node()      # velocity
            a = Node()      # accelerator
            # a rough solution but it works
            # need to change a lot to fit other situation
            for d in ds:
                if '!' in d:
                    state = 0
                elif '\'' in d:
                    state = 1
                elif '"' in d:
                    state = 2
                temp_array = []
                indicator = 0
                while True:
                    while d[0] in [',', '!', '\'', '"', ' ']:
                        d = d[1:]
                    if indicator >= len(d):
                        temp_array.append(d)
                        break
                    if d[indicator] not in [',', '!', '\'', '"', ' ', '-'] or indicator == 0:
                        indicator += 1
                        continue
                    temp_array.append(d[:indicator])
                    d = d[indicator:]
                    indicator = 0
                while len(temp_array) < 5:      # if no force data when onenote is set to ignore pen pressure
                    temp_array.append(0)
                temp_array = list(map(int, temp_array))
                if state == 0:
                    c.set(temp_array)
                elif state == 1:
                    v.set(temp_array)
                    c.set_with_v(v)
                elif state == 2:
                    a.set(temp_array)
                    v.set_with_v(a)
                    c.set_with_v(v)
                node = Node()
                node.copy(c)
                trace.add(node)
            self.traces.append(trace)

    def to_json(self):
        data = {'type': "excalidraw/clipboard", "elements": []}
        base_node = (self.traces[0].nodes[0].x, self.traces[0].nodes[0].y)
        for trace in self.traces:
            ele = self.init_element()
            ele["id"] = str(random.randint(1000000, 10000000))
            ele["x"] = (trace.nodes[0].x - base_node[0]) * self.overall_scale
            ele["y"] = (trace.nodes[0].y - base_node[1]) * self.overall_scale
            trace.normalize(self.overall_scale)

            ele["width"] = trace.size[0]
            ele["height"] = trace.size[1]
            ele["strokeColor"] = trace.brush.color
            ele["strokeWidth"] = trace.brush.height * self.pen_scale
            ele["opacity"] = int(trace.brush.transparency / 255.0 * 100)
            points = []
            pressures = []
            cnt = -1
            for node in trace.nodes:
                cnt += 1
                if cnt % self.sample_rate != 0 and cnt != len(trace.nodes) - 1:
                    continue
                points.append([node.x, node.y])
                pressures.append(node.force)
            ele["points"] = points
            # seems the "simulatePressure" in Excal json is not what I think
            if self.force_ignore_pressure:
                ele["pressures"] = [1 for x in pressures]
                # ele["simulatePressure"] = False
            else:
                ele["pressures"] = pressures
                # ele["simulatePressure"] = False if trace.brush.ignore_pressure == 1 else True
            data["elements"].append(ele)

        def round_floats(o):
            if isinstance(o, float):
                return round(o, 3)
            if isinstance(o, dict):
                return {k: round_floats(v) for k, v in o.items()}
            if isinstance(o, (list, tuple)):
                return [round_floats(x) for x in o]
            return o
        self.json = json.dumps(round_floats(data))

    def init_element(self):
        return {"id": "123", "type": "freedraw",
                "x": 0, "y": 0,
                "width": 0, "height": 0, "angle": 0,
                "strokeColor": "#000000", "backgroundColor": "transparent", "fillStyle": "hachure",
                "strokeWidth": 0.5, "strokeStyle": "solid", "roughness": None, "opacity": 100,
                "groupIds": [], "frameId": None, "index": "a1", "roundness": None,
                "seed": 1588510656, "version": 13, "versionNonce": 903747520,
                "isDeleted": False, "boundElements": None,
                "updated": 1718792668563, "link": None, "locked": False,
                "points": [], "pressures": [], "simulatePressure": False
                }

    def export(self):
        if self.export_to_file:
            with open(self.file_path, "w") as f:
                f.write(a.json)
            print(f"Export to {self.file_path} success.")
        else:
            wcb.OpenClipboard(None)
            wcb.SetClipboardText(self.json.encode('utf-8'))
            wcb.CloseClipboard()
            print("Export to clipboard success.")


class Node:
    def __init__(self):
        self.x = 0
        self.y = 0
        self.force = 0
        self.oa = 0
        self.oe = 0

    def set(self, data):
        self.x = data[0]
        self.y = data[1]
        self.force = data[2]
        self.oa = data[3]
        self.oe = data[4]

    def set_with_v(self, v):
        self.x += v.x
        self.y += v.y
        self.force += v.force
        self.oa += v.oa
        self.oe += v.oe

    def copy(self, n):
        self.x = n.x
        self.y = n.y
        self.force = n.force
        self.oa = n.oa
        self.oe = n.oe


class Trace:
    def __init__(self):
        self.nodes: list[Node] = []
        # self.context = None
        self.brush: Brush = None
        self.size = (0, 0)

    def add(self, node: Node):
        self.nodes.append(node)

    def calc_size(self):
        minx = 9999999
        maxx = -9999999
        miny = 9999999
        maxy = -9999999
        for node in self.nodes:
            if node.x < minx:
                minx = node.x
            if node.x > maxx:
                maxx = node.x
            if node.y < miny:
                miny = node.y
            if node.y > maxy:
                maxy = node.y
        self.size = (maxx-minx, maxy-miny)

    def normalize(self, scale):
        self.calc_size()
        basex = self.nodes[0].x
        basey = self.nodes[0].y
        for node in self.nodes:
            node.x = (node.x - basex) * scale
            node.y = (node.y - basey) * scale

            # magic number for I don't read force range in inkML
            # change this if you have issue (maybe older onenote version?)
            # find this in inkML in clipboard
            node.force /= 32767.0

            if self.brush.ignore_pressure == 1:
                node.force = 1


class Brush:
    def __init__(self):
        self.width = 0
        self.height = 0
        self.color = "#000000"
        self.ignore_pressure = 0
        self.transparency = 255


formats = {val: name for name, val in vars(wcb).items() if name.startswith('CF_')}


def format_name(fmt):
    if fmt in formats:
        return formats[fmt]
    try:
        return wcb.GetClipboardFormatName(fmt)
    except:
        return "unknown"


def handle_clipboard():
    wcb.OpenClipboard(None)
    fmt = 0
    data = None
    while True:
        fmt = wcb.EnumClipboardFormats(fmt)
        if fmt == 0:
            break
        if format_name(fmt) == "InkML Format":
            hg = wcb.GetClipboardDataHandle(fmt)
            data = wcb.GetGlobalMemory(hg)
            wcb.EmptyClipboard()
            # print(data)
            break
    wcb.CloseClipboard()
    return data


if __name__ == "__main__":
    while True:
        text = input("Copy Onenote ink and press enter.")
        data = handle_clipboard()
        if data is not None:
            a = InkParser(data)
            a.read_inkml()
            a.to_json()
            a.export()
        else:
            print("Could not detect Onenote ink data. Please copy again and retry.")
