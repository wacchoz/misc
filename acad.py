# -*- coding: utf-8 -*-

"""
AutoCADの作図

AcadRemoconが必要
"""

import math
import win32com.client

class Point(object):
    def __init__(self, x=0, y=0):
        self.x = x
        self.y = y
            
    def __add__(self, other):
        return Point( self.x + other.x, self.y + other.y)

    def __sub__(self, other):
        return Point( self.x - other.x, self.y - other.y)
        
    def __mul__(self, other):
        return Point( self.x * other, self.y * other )
    
    def __truediv__(self, other):
        return Point( self.x / other, self.y / other)
        
    def __neg__(self):
        return Point(-self.x, -self.y)
        
    def __pos__(self):
        return Point(self.x, self.y)
            
    def __str__(self):
        return "NONE " + str(self.x) + "," + str(self.y)
#        return str(self.x) + "," + str(self.y) 
            
    def mirror(self):
        return Point(-self.x, self.y)


# 線分pt1, pt2と線分pt3, pt4が交わるかどうかのチェック
# 交われば交点を返す
def intersect_segment(pt1, pt2, pt3, pt4):
    # 参考：http://www5d.biglobe.ne.jp/~tomoya03/shtml/algorithm/Intersection.htm
    if ((pt1.x - pt2.x) * (pt3.y - pt1.y) + (pt1.y - pt2.y) * (pt1.x - pt3.x)) \
        * ((pt1.x - pt2.x) * (pt4.y - pt1.y) + (pt1.y - pt2.y) * (pt1.x - pt4.x)) > 0 \
        or \
       ((pt3.x - pt4.x) * (pt1.y - pt3.y) + (pt3.y - pt4.y) * (pt3.x - pt1.x)) \
            * ((pt3.x - pt4.x) * (pt2.y - pt3.y) + (pt3.y - pt4.y) * (pt3.x - pt2.x)) > 0:

            return False, None
    else:
        # 交点を計算する
        dev = (pt2.y-pt1.y)*(pt4.x-pt3.x)-(pt2.x-pt1.x)*(pt4.y-pt3.y)
        if dev == 0:
            return False, None
            
        d1 = (pt3.y*pt4.x-pt3.x*pt4.y)
        d2 = (pt1.y*pt2.x-pt1.x*pt2.y) 
        result = Point()
        result.x = d1*(pt2.x-pt1.x) - d2*(pt4.x-pt3.x)
        result.x /= dev
        result.y = d1*(pt2.y-pt1.y) - d2*(pt4.y-pt3.y)
        result.y /= dev

        return True, result


# 直線pt1, pt2と線分pt3, pt4が交わるかどうかのチェック
# 交われば交点を返す
def intersect_line_and_segment(pt1, pt2, pt3, pt4):
    # 参考：http://www5d.biglobe.ne.jp/~tomoya03/shtml/algorithm/Intersection.htm
    if ((pt1.x - pt2.x) * (pt3.y - pt1.y) + (pt1.y - pt2.y) * (pt1.x - pt3.x)) \
        * ((pt1.x - pt2.x) * (pt4.y - pt1.y) + (pt1.y - pt2.y) * (pt1.x - pt4.x)) > 0:

        return False, None
    else:
        # 交点を計算する
        dev = (pt2.y-pt1.y)*(pt4.x-pt3.x)-(pt2.x-pt1.x)*(pt4.y-pt3.y)
        if dev == 0:
            return False, None
            
        d1 = (pt3.y*pt4.x-pt3.x*pt4.y)
        d2 = (pt1.y*pt2.x-pt1.x*pt2.y) 
        result = Point()
        result.x = d1*(pt2.x-pt1.x) - d2*(pt4.x-pt3.x)
        result.x /= dev
        result.y = d1*(pt2.y-pt1.y) - d2*(pt4.y-pt3.y)
        result.y /= dev
 
        return True, result

# 直線pt1, pt2と直線pt3, pt4が交わるかどうかのチェック
# 交われば交点を返す
def intersect_line(pt1, pt2, pt3, pt4):

    dev = (pt2.y-pt1.y)*(pt4.x-pt3.x)-(pt2.x-pt1.x)*(pt4.y-pt3.y)
    if dev == 0:
        return False, None
        
    d1 = (pt3.y*pt4.x-pt3.x*pt4.y)
    d2 = (pt1.y*pt2.x-pt1.x*pt2.y) 
    result = Point()
    result.x = d1*(pt2.x-pt1.x) - d2*(pt4.x-pt3.x)
    result.x /= dev
    result.y = d1*(pt2.y-pt1.y) - d2*(pt4.y-pt3.y)
    result.y /= dev
 
    return True, result



class Acad(object):
    def __init__(self):
        self._acad=win32com.client.Dispatch("AcadRemocon.Body")

        self.current_layer = "0"      
        
        self.direction = +1
        self.offset = Point(0,0)

    def acPostCommand(self, command):
        self._acad.acPostCommand(command)

    def line(self, x1, y1, x2, y2):
        if self.direction == 1:
            self._acad.acPostCommand("LINE " + str(x1+self.offset.x) + "," + str(y1+self.offset.y) + " " + str(x2+self.offset.x) + "," + str(y2+self.offset.y) + "^M^C^C")
        else:
            self._acad.acPostCommand("LINE " + str(-x1+self.offset.x) + "," + str(y1+self.offset.y) + " " + str(-x2+self.offset.x) + "," + str(y2+self.offset.y) + "^M^C^C")
       
    def linePt(self, pt1, pt2):
        if self.direction == 1:
            self._acad.acPostCommand("LINE " + str(pt1.x + self.offset.x) + "," + str(pt1.y+self.offset.y) + " " + str(pt2.x+self.offset.x) + "," + str(pt2.y+self.offset.y) + "^M^C^C")
        else:
            self._acad.acPostCommand("LINE " + str(-pt1.x + self.offset.x) + "," + str(pt1.y+self.offset.y) + " " + str(-pt2.x+self.offset.x) + "," + str(pt2.y+self.offset.y) + "^M^C^C")

    def linePt_rel(self, pt1, rel):
        self.linePt(pt1, pt1 + rel)

    def linePt_multi(self, pt_list):
        pt1 = pt_list[0]
        for pt in pt_list[1:]:
            self.linePt(pt1, pt)
            pt1 = pt

    def circle(self, x, y, r):
        if self.direction == 1:
            self._acad.acPostCommand("CIRCLE^M" + str(x+self.offset.x) + "," + str(y+self.offset.y) + "^M" + str(r) + "^M^C^C")
        else:
            self._acad.acPostCommand("CIRCLE^M" + str(-x+self.offset.x) + "," + str(y+self.offset.y) + "^M" + str(r) + "^M^C^C")

    def circlePt(self, pt, r):
        if self.direction == 1:
            self._acad.acPostCommand("CIRCLE^M" + str(pt.x+self.offset.x) + "," + str(pt.y+self.offset.y) + "^M" + str(r) + "^M^C^C")
        else:
            self._acad.acPostCommand("CIRCLE^M" + str(-pt.x+self.offset.x) + "," + str(pt.y+self.offset.y) + "^M" + str(r) + "^M^C^C")

    def arc(self, x, y, r, angle1, angle2, dimension=False, dimlayer="寸法"):
        if self.direction == 1:
            self._acad.acPostCommand ("ARC^MC^M" + str(x+self.offset.x) + "," + str(y+self.offset.y) + "^M^M@" + str(r) + "<" + str(angle1) + "^M^M@" + str(r) + "<" + str(angle2) + "^M^C^C")

            if dimension:
                layer = self.current_layer
                self.setLayer(dimlayer)

                mean_angle = angle1/2 + angle2/2
                self.radius_dimensionPt( Point(x,y) + Point(math.cos(math.radians(mean_angle)), math.sin(math.radians(mean_angle))) * r, Point(x,y) )

                self.setLayer(layer)

        else:
            self._acad.acPostCommand ("ARC^MC^M" + str(-x+self.offset.x) + "," + str(y+self.offset.y) + "^M^M@" + str(r) + "<" + str(180-angle2) + "^M^M@" + str(r) + "<" + str(180-angle1) + "^M^C^C")

            if dimension:
                layer = self.current_layer
                self.setLayer(dimlayer)

                mean_angle = (180-angle1)/2 + (180-angle2)/2
                self.radius_dimensionPt( Point(-x,y) + Point(math.cos(math.radians(mean_angle)), math.sin(math.radians(mean_angle))) * r, Point(-x,y) )

                self.setLayer(layer)

    def arcPt(self, pt, r, angle1, angle2, dimension=False, dimlayer="寸法"):
        if self.direction == 1:
            self._acad.acPostCommand ("ARC^MC^M" + str(pt.x+self.offset.x) + "," + str(pt.y+self.offset.y) + "^M^M@" + str(r) + "<" + str(angle1) + "^M^M@" + str(r) + "<" + str(angle2) + "^M^C^C")

        else:
            self._acad.acPostCommand ("ARC^MC^M" + str(-pt.x+self.offset.x) + "," + str(pt.y+self.offset.y) + "^M^M@" + str(r) + "<" + str(180-angle2) + "^M^M@" + str(r) + "<" + str(180-angle1) + "^M^C^C")

        if dimension:
            layer = self.current_layer
            self.setLayer(dimlayer)

            mean_angle = angle1/2 + angle2/2
            self.radius_dimensionPt( pt + Point(math.cos(math.radians(mean_angle)), math.sin(math.radians(mean_angle))) * r, pt + Point(math.cos(math.radians(mean_angle)), math.sin(math.radians(mean_angle))) * (r/2))

            self.setLayer(layer)

    def dxfCopy(self, filename, origin=Point(0,0), direction=+1):
        count = 0 # dummy
        ExtractArray=[[]] # dummy
        
        # 10: startX, 20: startY, 11: endX, 21: endY
        ret = self._acad.DxfExtract(count, ExtractArray, "ENTITIES", "", "LINE", "10|20|11|21", filename)
        if ret[0]:
            startX = ret[2][5][1:]
            startY = ret[2][6][1:]
            endX = ret[2][7][1:]
            endY = ret[2][8][1:]
            
            for i in range(len(startX)):
#                self.line( float(startX[i]), float(startY[i]), float(endX[i]), float(endY[i]), self.offset + origin, self.direction )
                if direction == +1:
                    x1 = float(startX[i]) + origin.x
                    y1 = float(startY[i]) + origin.y
                    x2 = float(endX[i]) + origin.x
                    y2 = float(endY[i]) + origin.y
                else:
                    x1 = -float(startX[i]) + origin.x
                    y1 = float(startY[i]) + origin.y
                    x2 = -float(endX[i]) + origin.x
                    y2 = float(endY[i]) + origin.y

                if self.direction == 1:
                    self._acad.acPostCommand("LINE " + str(x1+self.offset.x) + "," + str(y1+self.offset.y) + " " + str(x2+self.offset.x) + "," + str(y2+self.offset.y) + "^M^C^C")
                else:
                    self._acad.acPostCommand("LINE " + str(-x1+self.offset.x) + "," + str(y1+self.offset.y) + " " + str(-x2+self.offset.x) + "," + str(y2+self.offset.y) + "^M^C^C")

        # 10: centerX, 20: centerY, 40: radius, 50:startAngle, 51: endAngle
        ret = self._acad.DxfExtract(count, ExtractArray, "ENTITIES", "", "ARC", "10|20|40|50|51", filename)
        if ret[0]:        
            centerX = ret[2][6][1:]
            centerY = ret[2][7][1:]
            radius = ret[2][8][1:]
            start_angle = ret[2][9][1:]
            end_angle = ret[2][10][1:]
    
            for i in range(len(centerX)):
#                self.arc(float(centerX[i]), float(centerY[i]), float(radius[i]), float(start_angle[i]), float(end_angle[i]), self.offset + origin, self.direction)
                if direction == +1:
                    x = float(centerX[i]) + origin.x
                    y = float(centerY[i]) + origin.y
                else:
                    x = -float(centerX[i]) + origin.x
                    y = float(centerY[i]) + origin.y
                    
                r = float(radius[i])
                angle1 = float(start_angle[i])
                angle2 = float(end_angle[i])

                if direction == 1:
                    if self.direction == 1:
                        self._acad.acPostCommand ("ARC^MC^M" + str(x+self.offset.x) + "," + str(y+self.offset.y) + "^M^M@" + str(r) + "<" + str(angle1) + "^M^M@" + str(r) + "<" + str(angle2) + "^M^C^C")
                    else:
                        self._acad.acPostCommand ("ARC^MC^M" + str(-x+self.offset.x) + "," + str(y+self.offset.y) + "^M^M@" + str(r) + "<" + str(180-angle2) + "^M^M@" + str(r) + "<" + str(180-angle1) + "^M^C^C")
                else:
                    if self.direction == 1:
                        self._acad.acPostCommand ("ARC^MC^M" + str(x+self.offset.x) + "," + str(y+self.offset.y) + "^M^M@" + str(r) + "<" + str(180-angle2) + "^M^M@" + str(r) + "<" + str(180-angle1) + "^M^C^C")
                        pass
                    else:
                        self._acad.acPostCommand ("ARC^MC^M" + str(-x+self.offset.x) + "," + str(y+self.offset.y) + "^M^M@" + str(r) + "<" + str(angle1) + "^M^M@" + str(r) + "<" + str(angle2) + "^M^C^C")


    def makeLayer(self, name, color):
        self._acad.acPostCommand ("-layer^MN^M" + name + "^M^C")
        self._acad.acPostCommand ("-layer^MC^M" + str(color) + "^M" + name + "^M^C")
       
    def setLayer(self, name):
        self._acad.acPostCommand ("-layer^MS^M" + name + "^M^C")

        self.current_layer = name

    def linear_VdimensionPt(self, pt1, pt2, pt3=None):
        if pt3 == None:
            pt3 = (pt1 + pt2) * 0.5
        else:
            if pt3.x == None:
                pt3.x = (pt1.x + pt2.x) * 0.5
            if pt3.y == None:
                pt3.y = (pt1.y + pt2.y) * 0.5
            
        if self.direction ==1:
            self._acad.acPostCommand("DIMLINEAR^M" + str(pt1 + self.offset) + "^M" \
                                + str(pt2 + self.offset) + "^MV^M" \
                                + str(pt3 + self.offset) + "^M^C^C")
        else:
            self._acad.acPostCommand("DIMLINEAR^M" + str(pt1.mirror() + self.offset) + "^M" \
                                + str(pt2.mirror() + self.offset) + "^MV^M" \
                                + str(pt3.mirror() + self.offset) + "^M^C^C")


    def linear_HdimensionPt(self, pt1, pt2, pt3=None):
        if pt3 == None:
            pt3 = (pt1 + pt2) * 0.5
        else:
            if pt3.x == None:
                pt3.x = (pt1.x + pt2.x) * 0.5
            if pt3.y == None:
                pt3.y = (pt1.y + pt2.y) * 0.5

        if self.direction ==1:
            self._acad.acPostCommand("DIMLINEAR^M" + str(pt1 + self.offset) + "^M" \
                                + str(pt2 + self.offset) + "^MH^M" \
                                + str(pt3 + self.offset) + "^M^C^C")
        else:
            self._acad.acPostCommand("DIMLINEAR^M" + str(pt1.mirror() + self.offset) + "^M" \
                                + str(pt2.mirror() + self.offset) + "^MH^M" \
                                + str(pt3.mirror() + self.offset) + "^M^C^C")


    def radius_dimensionPt(self, pt1, pt2=None):
        if pt2 is None:
            pt2 = pt1
            
        if self.direction == 1:
            self._acad.acPostCommand("DIMRADIUS^M" + str(pt1 + self.offset) + "^M" + str(pt2 + self.offset) + "^M^C^C" )
        else:
            self._acad.acPostCommand("DIMRADIUS^M" + str(pt1.mirror() + self.offset) + "^M" + str(pt2.mirror() + self.offset) + "^M^C^C" )


    def add_text(self, text, pt, height=3.5, align="M", angle=0):
        if self.direction == 1:
            x_point = pt.x + self.offset.x
            y_point = pt.y + self.offset.y
        else:
            x_point = -pt.x + self.offset.x
            y_point = pt.y + self.offset.y

        self._acad.acPostCommand("-TEXT^MJ^M" + align + "^M" + str(Point(x_point, y_point)) + "^M" + str(height) + "^M" + str(angle) + "^M" + text + "^M^C^C")


    # pt1, pt2, pt3を線分で結びながら、radiusでフィレット
    def two_line_with_fillet(self, pt1, pt2, pt3, radius, dimension=False, dimlayer="寸法"):
        
        # pt1からpt2の角度(0-2pi)
        vec12 = pt2 - pt1
        if vec12.x == 0 and vec12.y == 0:
            self.linePt(pt1,pt3)
            return

        alpha = math.acos( vec12.x / math.sqrt(vec12.x**2 + vec12.y**2) )
        if vec12.y < 0:
            alpha = 2*math.pi - alpha
        
        # pt2からpt3の角度(0-2pi)
        vec23 = pt3 - pt2
        if vec23.x == 0 and vec23.y == 0:
            self.linePt(pt1,pt3)
            return

        beta = math.acos( vec23.x / math.sqrt(vec23.x**2 + vec23.y**2) )
        if vec23.y < 0:
            beta = 2*math.pi - beta

        # 完全に一直線になっている場合
        if alpha == beta:
            self.linePt(pt1, pt3)
            return

        if 0 < beta-alpha < math.pi or -2*math.pi < beta-alpha < -math.pi:
            theta = math.pi / 2 + (alpha+beta)/2
            center = pt2 + Point(math.cos(theta), math.sin(theta)) * (radius / math.cos( (beta-alpha)/2 ))

            angle1 = math.pi*3/2+alpha
            angle2 = math.pi*3/2+beta
            if angle1 >= 2*math.pi:
                angle1 -= 2*math.pi
            if angle2 >= 2*math.pi:
                angle2 -= 2*math.pi
    
            self.linePt(pt1, pt2 + Point(math.cos(math.pi + alpha), math.sin(math.pi + alpha)) * (radius * math.tan((beta-alpha)/2)) )
            self.linePt(pt2 - Point(math.cos(math.pi + beta), math.sin(math.pi + beta)) * (radius * math.tan((beta-alpha)/2)), pt3)
            self.arcPt(center, radius, math.degrees(angle1), math.degrees(angle2), dimension=dimension, dimlayer=dimlayer )

        else:
            theta = (alpha+beta)/2 + math.pi/2
            center = pt2 - Point(math.cos(theta), math.sin(theta)) * (radius / math.cos( (beta-alpha)/2 ))

            angle1 = math.pi/2+beta
            angle2 = math.pi/2+alpha
            if angle1 >= 2*math.pi:
                angle1 -= 2*math.pi
            if angle2 >= 2*math.pi:
                angle2 -= 2*math.pi
    
            self.linePt(pt1, pt2 + Point(math.cos(alpha), math.sin(alpha)) * (radius * math.tan((beta-alpha)/2)) )
            self.linePt(pt2 - Point(math.cos(beta), math.sin(beta)) * (radius * math.tan((beta-alpha)/2)), pt3)
            self.arcPt(center, radius, math.degrees(angle1), math.degrees(angle2), dimension=dimension, dimlayer=dimlayer )


    def zoom(self):
        self._acad.acPostCommand("_zoom^ME^M")


    def zoomPt(self, pt1, pt2):
        self._acad.acPostCommand("_zoom^MW^M" + str(pt1) + "^M" + str(pt2) + " ^M^C")
        

if __name__ == '__main__':

    acad = Acad()
    acad.makeLayer("寸法", "gray")

    acad.linear_HdimensionPt(Point(0,0),Point(100,100))

    acad.two_line_with_fillet(Point(0,0), Point(100,0), Point(100,100), 30)
    acad.two_line_with_fillet(Point(0,0), Point(100,0), Point(70,-70), 30)
    acad.two_line_with_fillet(Point(0,0), Point(0,100), Point(-100,100), 30)
    acad.two_line_with_fillet(Point(0,0), Point(0,100), Point(100,100), 30)
    acad.two_line_with_fillet(Point(0,0), Point(-100,0), Point(-100,-100), 30)
    acad.two_line_with_fillet(Point(0,0), Point(-100,0), Point(-100,100), 30)
    acad.two_line_with_fillet(Point(0,0), Point(0,-100), Point(100,-100), 30)
    acad.two_line_with_fillet(Point(100,100), Point(100,200), Point(130,210), 30)
    acad.two_line_with_fillet(Point(100,100), Point(100,200), Point(90,210), 30)
    acad.two_line_with_fillet(Point(100,100), Point(10,10), Point(110,0), 30)
    acad.two_line_with_fillet(Point(100,100), Point(100,200), Point(90,210), 30)

    acad.zoomPt(Point(0,0),Point(3600,1000))
    for angle in range(0,360,30):
        acad.arcPt(Point(angle*10,0),100,0,angle, dimension=True)

    acad.offset = Point(3600,1000)  # オフセット位置を原点とする
    acad.direction = -1             # 逆向きに
    for angle in range(0,360,30):
        acad.arcPt(Point(angle*10,0),100,0,angle, dimension=True)
