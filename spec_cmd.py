# -*- coding: utf8 -*-

import Rhino
import Rhino.Geometry as g
import rhinoscriptsyntax as rs
import codecs
import sys
import System
import Rhino.UI
import Eto.Drawing as drawing
import Eto.Forms as forms
import Rhino.UI
import scriptcontext
import math
import webbrowser

__commandname__ = "spec"

LayersEffect = 0
DIR = {"x":(0,1,2),
"y":(1,0,2),
"z":(2,0,1)}
AXES = {0:Rhino.Geometry.Vector3d.XAxis,
1:Rhino.Geometry.Vector3d.YAxis,
2:Rhino.Geometry.Vector3d.ZAxis
}

def computeDims(obj):
    objRef = Rhino.DocObjects.ObjRef(obj)
    brep = objRef.Geometry()    
    edges = list(brep.Edges)
    edges.sort(key=lambda edge: edge.GetLength(), reverse=True)
    edge = edges[0]
    adFaces = list(edge.AdjacentFaces())
    faces = list(brep.Faces)
    adFaces = [faces[i] for i in adFaces]
    adFaces.sort(key=lambda face:face.GetSurfaceSize()[1]*face.GetSurfaceSize()[2],reverse=True)
    face = adFaces[0]
    faceNor = face.NormalAt(0,0)
    origin = edge.PointAtStart
    xPoint = edge.PointAtEnd
    xVec = Rhino.Geometry.Vector3d(xPoint-origin)
    xVec.Unitize()
    yVec = Rhino.Geometry.Vector3d.CrossProduct(faceNor, xVec)
    yVec.Unitize()
    plane = Rhino.Geometry.Plane(origin, xVec, yVec)
    surf = Rhino.Geometry.PlaneSurface(plane, Rhino.Geometry.Interval(0,1),Rhino.Geometry.Interval(0,1))
    scriptcontext.doc.Objects.AddSurface(surf)
    verts = list(brep.Vertices)
    worldPlane = g.Plane.WorldXY
    tMatrix = g.Transform.ChangeBasis(worldPlane, plane)
    map(lambda vert:vert.Transform(tMatrix),verts)
    verts = list(map(lambda vert:vert.Location,verts))
    verts.sort(key=lambda vert:vert.X)
    length = verts[len(verts)-1].X - verts[0].X
    verts.sort(key=lambda vert:vert.Y)
    width = verts[len(verts)-1].Y - verts[0].Y
    verts.sort(key=lambda vert:vert.Z)
    thickness = verts[len(verts)-1].Z - verts[0].Z
    return [length,width,thickness]

class attr:
    def __init__(self, name, isEditable, isOn, valType, exportPos):
        self.name = name
        self.isEditable = isEditable
        self.isOn = isOn
        self.valType = valType
        self.exportPos = exportPos

ATTRS = (attr("N",True,True,int,1),
attr("Name",True,True,str,0),
attr("Length",False,True,int,2),
attr("Width",False,True,int,3),
attr("Thickness",False,True,int,4),
attr("Texture",True,True,str,-1),
attr("Edge_L_0.4",True,True,str,6),
attr("Edge_W_0.4",True,True,str,7),
attr("Edge_L",True,True,str,8),
attr("Edge_W",True,True,str,9),
attr("Layer",False,True,str,10),
attr("Comment",True,True,str,11)
)

QUANT_FIELD = attr("number",False,True,int,5)

QUANTITY_COLUMN_POS = 5

MAX_DIM_LEN = 2780
MAX_DIM_WIDTH = 2050

DIM_PRECISION = rs.GetDocumentUserText ("easy_cut_precision")
if DIM_PRECISION==None:
    DIM_PRECISION = 0
else:
    try:
        DIM_PRECISION = int(DIM_PRECISION)
    except:
        DIM_PRECISION = 0

class Specification:
    QUANTITY_COLUMN_POS = 5
    def __init__(self, attrs, details):
        self.details = [] #{}
        idSet = set(range(len(details)))
        self.table = []
        while len(idSet)>0:
            curId = idSet.pop()
            curDet = details[curId]
            curDetails = [curDet]
            idLst = list(idSet)
            for id in idLst:
                if curDet == details[id]:
                    curDetails.append(details[id])
                    idSet.remove(id)
            self.details.append(curDetails)
        self.buildTable()
    def buildTable(self):
        self.table = []
        for details in self.details:
            specRow = list(details[0].getParams())
            specRow.insert(self.QUANTITY_COLUMN_POS, len(details))
            self.table.append(specRow)
        self.table, self.details = (list(t) for t in zip(*sorted(zip(self.table, self.details), key=lambda x: (int(x[0][0]) if int(x[0][0])>0 else 9999, x[0][10], -float(x[0][4])))))
    def autoNum(self):
        i = 1
        for tableItem in self.table:
            for detail in self.details[i-1]:
                detail.N = i
                detail.refresh()
            tableItem[0] = i
            i += 1
            


class SpecDialog(forms.Dialog[bool]):

    def __init__(self):
        self.Title = "EasyCut"
        self.selected = []
        self.edgeshighlightMode = -1
        self.m_gridview = forms.GridView()
        self.m_gridview.ShowHeader = True
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        num=0
        for idx in range(len(ATTRS)):
            if num == QUANTITY_COLUMN_POS:
                attr = QUANT_FIELD
                column = forms.GridColumn()
                column.HeaderText = attr.name
                column.Editable = attr.isEditable
                column.DataCell = forms.TextBoxCell(num)
                self.m_gridview.Columns.Add(column)
                num+=1
            attr = ATTRS[idx]
            if attr.isOn:

                column = forms.GridColumn()
                column.HeaderText = attr.name
                column.Editable = attr.isEditable
                column.DataCell = forms.TextBoxCell(num)
                self.m_gridview.Columns.Add(column)
                num+=1
                
        self.precision_dropdownlist = forms.DropDown()
        self.precision_dropdownlist.DataStore = ['1', '0.1', '0.01', '0.001', '0.0001', '0.00001']
        self.precision_dropdownlist.SelectedIndex = DIM_PRECISION
        self.precision_label = forms.Label(Text = "Dimensions precision " )
        self.precision_dropdownlist.DropDownClosed += self.changePrecisionVal
        
        layout0 = forms.TableLayout()
        cell = forms.TableCell(self.precision_label,scaleWidth=False)
        cell.ScaleWidth = False
        cell2 = forms.TableCell(self.precision_dropdownlist,scaleWidth=False)
        cell2.ScaleWidth = False
        row = forms.TableRow(None, cell, cell2)
        row.ScaleHeight = False
        layout0.Rows.Add(row)
        
        
        self.m_gridview.CellClick += self.gridClick
        self.m_gridview.SelectionChanged += self.gridSelChanged
        self.m_gridview.CellEdited += self.gridEdited
        self.m_gridview.CellFormatting += self.OnCellFormatting
        self.buttonAutoNum = forms.Button(self.buttonAutoNumClick)
        self.buttonAutoNum.Text = "Auto number"
        self.button = forms.Button(self.buttonClick)
        self.button.Text = "Export"
        layout = forms.TableLayout()
        layout.Spacing = drawing.Size(5, 5)
        cell = forms.TableCell(self.m_gridview)
        row = forms.TableRow(cell)
        row.ScaleHeight = True
        layout.Rows.Add(layout0)
        layout.Rows.Add(row)
        
        layout2 = forms.TableLayout()
        layout2.Spacing = drawing.Size(5, 5)
        
        cell = forms.TableCell(self.button,True)
        cell2 = forms.TableCell(self.buttonAutoNum,True)
        row = forms.TableRow([cell,cell2])
        
        layout2.Rows.Add(row)
        layout.Rows.Add(layout2)
        layout2 = forms.TableLayout()
        layout2.Spacing = drawing.Size(5, 5)
        self.m_linkbutton = forms.LinkButton(Text = 'Easycut')
        self.m_linkbutton.Click += self.OnLinkButtonClick
        self.m_donatebutton = forms.LinkButton(Text = 'Donate',Style = "right-align")
        self.m_donatebutton.Click += self.OnDonateButtonClick
        cell = forms.TableCell(self.m_linkbutton,True)
        cell2 = forms.TableCell(None,False)
        cell3 = forms.TableCell(self.m_donatebutton,False)
        row = forms.TableRow([cell,cell2,cell3])
        layout2.Rows.Add(row)
        
        layout.Rows.Add(layout2)
        
        self.Content = layout
        
    def changePrecisionVal(self, sender, e):
        global DIM_PRECISION
        DIM_PRECISION = sender.SelectedIndex
        rs.SetDocumentUserText ("easy_cut_precision", str(DIM_PRECISION))
        specification = self.spec.details
        for row in specification:
            for det in row:
                det.setDims()

        
        self.rebuild()

    
    def OnCellFormatting(self, sender, e):
        if float(self.spec.table[e.Row][2]) > MAX_DIM_LEN:
            e.BackgroundColor = drawing.Colors.Red 
    def setData(self, spec):
        self.spec = spec
        self.m_gridview.DataStore = spec.table

    def rebuild(self):
        self.spec.buildTable()
        #print self.spec.table
        self.m_gridview.DataStore = self.spec.table
    def setObjs(self, dic):
        self.objs = dic
    def unselectAll(self):
        for det in self.selected:
            det.unselect()
        self.selected = []
    def gridClick(self,sender,e):
        row = e.Row
        col = e.Column
        dets = self.spec.details[row]
        self.edgeshighlightMode = -1
        if col >= 7 and col <= 10:
            self.edgeshighlightMode = (col - 7)%2
        self.unselectAll()
        if self.edgeshighlightMode>=0:
            for det in dets:
                det.getEdges(AXES[det.lcs[self.edgeshighlightMode]])
        else:
            for det in dets:
                det.select()
            self.selected += dets
        scriptcontext.doc.Views.Redraw()
        
    def gridSelChanged(self,sender,e):
        self.unselectAll()
        dets = []
        for row in sender.SelectedRows:
            dets += self.spec.details[row]
        for detLst in self.spec.details:
            for det in detLst:
                det.unhighlighAll()
        scriptcontext.doc.Views.Redraw()
        if self.edgeshighlightMode >=0:
            for det in dets:
                det.getEdges(AXES[det.lcs[self.edgeshighlightMode]])
        else:
           for det in dets:
               det.select()
           self.selected += dets
        scriptcontext.doc.Views.Redraw()


    def gridEdited(self, sender, e):
        row = e.Row
        col = e.Column
        dets = self.spec.details[row]
        cols = self.m_gridview.Columns
        attr = cols[col].HeaderText
        val = self.m_gridview.DataStore[row][col]
        if attr == "N":
            try:
                int(val)
            except:
                val=dets[0].N
                self.m_gridview.DataStore[row][col] = dets[0].N
                rs.MessageBox("number can be only an integer")
        for det in dets:
            setattr(det, attr, val)
            det.refresh()
        self.rebuild()
    def buttonClick(self, sender, e):
        data = []
        nameRow = []
        cols = [i.HeaderText for i in self.m_gridview.Columns]
        i = 0
        num=0
        curRow = []
        while i < (len(cols)):
            if i == QUANT_FIELD.exportPos:
                exportPos = QUANT_FIELD.exportPos
                if exportPos>0:
                    nameRow.append((exportPos,cols[i]))
                i += 1
            else:
                exportPos = ATTRS[num].exportPos
                if exportPos>=0:
                    nameRow.append((exportPos,cols[i]))
                num+=1
                i += 1
        nameRow = [i[1] for i in list(sorted(nameRow, key = lambda x: x[0]))]
        data.append(nameRow)
                
        for dataRow in self.m_gridview.DataStore:
            i = 0
            num=0
            curRow = []
            while i < (len(dataRow)):
                if i == QUANT_FIELD.exportPos:
                    exportPos = QUANT_FIELD.exportPos
                    if exportPos>0:
                        curRow.append((exportPos,dataRow[i]))
                    i += 1
                else:
                    exportPos = ATTRS[num].exportPos
                    if exportPos>=0:
                        curRow.append((exportPos,dataRow[i]))
                    num+=1
                    i += 1

            curRow = [i[1] for i in list(sorted(curRow, key = lambda x: x[0]))]
            data.append(curRow)

        ExportObjBBData(data)

    def buttonAutoNumClick(self, sender, e):
        self.spec.autoNum()
        self.rebuild()

    def OnLinkButtonClick(self, sender, e):
        webbrowser.open("https://easycut3d.online")
        
    def OnDonateButtonClick(self, sender, e):
        webbrowser.open("http://easycut3d.online/donate.html")

class Detail:
    def __init__(self, obj):
        self.obj = obj
        self.selected = 0
        self.id = obj
        attrs = rs.GetUserText (obj)
        attrs = [i.name for i in ATTRS if i.isEditable]
        for attr in attrs:
            setattr(self, attr, rs.GetUserText (str(obj), attr))
        self.setDims()
        objRef = Rhino.DocObjects.ObjRef(self.id)
        obj = objRef.Object()
        index = obj.Attributes.LayerIndex
        self.Layer = scriptcontext.doc.Layers[index].Name
    def setDims(self):        
#        bb=rs.BoundingBox(self.obj)
#        if bb:
#            x=round(bb[1].X-bb[0].X,DIM_PRECISION)
#            y=round(bb[3].Y-bb[0].Y,DIM_PRECISION)
#            z=round(bb[4].Z-bb[0].Z,DIM_PRECISION)            
#        else:
#            x=0 ; y=0 ; z=0
#        self.dims = [x,y,z]
        self.dims = computeDims(self.obj)
        texture = self.Texture
        if texture in DIR:
           dir = DIR[texture]
           self.Texture = texture
           self.lcs = dir
        else:
            dir = [0,1,2]
            self.lcs = [i[0] for i in sorted(list(enumerate(self.dims)), key = lambda x:x[1])]
            self.lcs.reverse()
            self.dims.sort()
            self.dims.reverse()
            alert = "texture missing"
        x,y,z = self.dims[dir[0]],self.dims[dir[1]],self.dims[dir[2]]
        if z>y:
            self.dims = [x,z,y]
            self.lcs = [self.lcs[0],self.lcs[2],self.lcs[1]]
        else:
            self.dims = [x,y,z]
        x = self.dims[0]
        y = self.dims[1]
        z = self.dims[2]
        self.Length = format(self.dims[0],'.' + str(DIM_PRECISION) + 'f')
        self.Width = format(self.dims[1],'.' + str(DIM_PRECISION) + 'f')
        self.Thickness = format(self.dims[2],'.' + str(DIM_PRECISION) + 'f')
    def select(self):
        self.selected = 1
        rs.SelectObject(self.id)
    def unselect(self):
        if self.selected:
            rs.UnselectObject(self.id)
    def getParams(self):
        params = []
        for attr in ATTRS:
            name = attr.name
            params.append(getattr(self, name))
        return params
    def highlightEdges(self, dir):
        pass
    def refresh(self):
        for attr in ATTRS:
            name = attr.name
            rs.SetUserText (str(self.id), name, getattr(self, name))
        self.setDims()
        for attr in ATTRS:
            name = attr.name
            rs.SetUserText (str(self.id), name, getattr(self, name))

    def getEdges(self, dirVec):
        id = self.id
        nor = AXES[self.lcs[2]]
        objRef = Rhino.DocObjects.ObjRef(id)
        brep = objRef.Geometry()
        obj = objRef.Object()
        faces = brep.Faces#doesn't work for extrusions
        edges = []
        for face in faces:
            obj.HighlightSubObject(face.ComponentIndex(), False)
            faceNor = face.NormalAt(0,0)
            cond1 = (round(faceNor*nor)==0) #perpendicular
            cond2 = (round(faceNor*dirVec)==0) #perpendicular
            if cond1 and cond2:               
               obj.HighlightSubObject(face.ComponentIndex(), True)
               scriptcontext.doc.Views.Redraw()
    def unhighlighAll(self):
        id = self.id
        objRef = Rhino.DocObjects.ObjRef(id)
        brep = objRef.Geometry()
        obj = objRef.Object()
        faces = brep.Faces#doesn't work for extrusions
        for face in faces:
            obj.HighlightSubObject(face.ComponentIndex(), False)
    def __eq__(self, obj):
        return (self.getParams() == obj.getParams())

        
def makeDetail(keys, objs):
    if not objs: return
    for obj in objs:
        for key in keys:
            cur_keys = rs.GetUserText (obj)
            if key not in cur_keys:
                rs.SetUserText (obj, key, "0", False)
            else:
                val = rs.GetUserText (obj, key)
                if len(val)<1 or val==" ":
                    rs.SetUserText (obj, key, "0", False)
def main():
    msg="Select polysurface objects to export data"
    objs = rs.GetObjects(msg,16,preselect=True)
    if not objs: return
    extrusions = []
    rs.UnselectAllObjects()
    for obj in objs:
        if rs.ObjectType(obj)==1073741824:
            extrusions.append(obj)
    if len(extrusions)>0:
        rs.SelectObjects(extrusions)
        resp = rs.MessageBox ("Selected objects will be converted to polysurfaces.\n Press OK to continue, Cancel to abort",1)
        if resp == 1:
            rs.Command("ConvertExtrusion _Enter")
            rs.UnselectAllObjects()
        else:
            return
    keys = [i.name for i in ATTRS if i.isEditable]
    makeDetail(keys, objs)
    details = []
    spec = dict([])
    ids = dict([])
    for obj in objs:
        detail = Detail(obj)
        details.append(detail)
    spec = Specification([], details)
    dialog = SpecDialog()
    dialog.setData(spec)
    rs.UnselectAllObjects()
    Rhino.UI.EtoExtensions.ShowSemiModal(dialog, Rhino.RhinoDoc.ActiveDoc, Rhino.UI.RhinoEtoApp.MainWindow)

def ExportObjBBData(data):
    filter = "CSV File (*.csv)|*.csv||"
    filename = rs.SaveFileName("Save data file as", filter)
    if not filename: return
    file = codecs.open(filename, encoding='utf-8', mode='w+')
    s = ""
    for itm in data:
        itm = [ i if isinstance(i, str) else str(i) for i in itm]
        s += ",".join(itm)
        s += "\n"
    file.write(s)
    file.close()


def RunCommand( is_interactive ):
    main()
    return 0

RunCommand(True)