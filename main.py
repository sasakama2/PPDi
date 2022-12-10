from kivy.config import Config

Config.set('graphics', 'fullscreen', '0')
Config.set('graphics', 'width', '670')
Config.set('graphics', 'height', '480')
Config.write()

from kivy.app import App
from kivy.core.text import DEFAULT_FONT, Label, LabelBase
from kivy.core.window import Window
from kivy.graphics import Color, Line
from kivy.lang import Builder
from kivy.properties import ObjectProperty, StringProperty,BooleanProperty
from kivy.resources import resource_add_path
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import (Screen, ScreenManager,
                                    ScreenManagerException)
from kivy.uix.widget import Widget
from kivy.utils import get_color_from_hex
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Pt

import pandas as pd
import numpy


resource_add_path('C:\Windows\Fonts')
LabelBase.register(DEFAULT_FONT, 'yuminl.ttf')
file_1='aaaaa'
file_2='aaaaa'
sm = ScreenManager()
file2=''
prs=None
sld1=None
sent='保存しました\n'



class diagnosis(Screen):
    def __init__(self, **kwargs):
        super(diagnosis, self).__init__(**kwargs)
        pass
    
    check=BooleanProperty(False)

    def dia(self,sent):
        ita=0
        bol=0
        und=0
        lef=0
                #見やすいstyle_id（４行目）
        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                if not shape.has_table:
                    continue
                tablee = shape.table
                tbl=tablee._graphic_frame.element.graphic.graphicData.tbl
                style_id =tbl[0][-1].text
                #print(style_id)#表のstyle_idを取得（確認用）
                
                new_style_id1='{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}'
                new_style_id2='{69012ECD-51FC-41F1-AA8D-1B2483CD663E}'
                new_style_id3='{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}'
                new_style_id4='{F2DE63D5-997A-4646-A377-4702673A728D}'
                new_style_id5='{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}'
                new_style_id6='{5A111915-BE36-4E01-A7E5-04B1672EAD32}'
                new_style_id7='{912C8C85-51F0-491E-9774-3900AFEF0FD7}'
                #見やすいstyle_id（４行目）
                
                if '{2D5ABB26-0587-4C30-8999-92F81FD0307C}' in style_id or '{5940675A-B579-460E-94D1-54222C63F5DA}'\
                    in style_id or '{616DA210-FB5B-4158-B5E0-FEB733F419BA}' in style_id or '{D7AC3CCA-C797-4891-BE02-D94E43425B78}'\
                        in style_id or '{E8034E78-7F5D-4C2E-B375-FC64B27BC917}' in style_id:
                            tbl[0][-1].text=new_style_id1
                
                if '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}' in style_id or '{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}'\
                    in style_id or '{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}' in style_id or '{69CF1AB2-1976-4502-BF36-3FF5EA218861}'\
                        in style_id or '{125E5076-3810-47DD-B79F-674D7AD40C01}' in style_id:
                            tbl[0][-1].text=new_style_id2
                
                if '{284E427A-3D55-4303-BF80-6455036E1DE7}' in style_id or '{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}'\
                    in style_id or '{5DA37D80-6434-44D0-A028-1B22A696006F}' in style_id or '{8A107856-5554-42FB-B03E-39F5DBC370BA}'\
                        in style_id or '{37CE84F3-28C3-443E-9E96-99CF82512B78}' in style_id:
                            tbl[0][-1].text=new_style_id3
                
                if '{69C7853C-536D-4A76-A0AE-DD22124D55A5}' in style_id or '{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}'\
                    in style_id or '{8799B23B-EC83-4686-B30A-512413B5E67A}' in style_id or '{0505E3EF-67EA-436B-97B2-0124C06EBD24}'\
                        in style_id or '{D03447BB-5D67-496B-8E87-E561075AD55C}' in style_id:
                            tbl[0][-1].text=new_style_id4
                
                if '{775DCB02-9BB8-47FD-8907-85C794F793BA}' in style_id or '{E269D01E-BC32-4049-B463-5C60D7B0CCD2}'\
                    in style_id or '{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}' in style_id or '{C4B1156A-380E-4F78-BDF5-A606A8083BF9}'\
                        in style_id or '{E929F9F4-4A8F-4326-A1B4-22849713DDAB}' in style_id:
                            tbl[0][-1].text=new_style_id5
                
                if '{35758FB7-9AC5-4552-8A53-C91805E547FA}' in style_id or '{327F97BB-C833-4FB7-BDE5-3F7075034690}'\
                    in style_id or '{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}' in style_id or '{22838BEF-8BB2-4498-84A7-C5851F593DF1}'\
                        in style_id or '{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}' in style_id:
                            tbl[0][-1].text=new_style_id6
                
                if '{08FB837D-C827-4EFA-A057-4D05807E0F7C}' in style_id or '{638B1855-1B75-4FBE-930C-398BA8C253C6}'\
                    in style_id or '{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}' in style_id or '{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}'\
                        in style_id or '{AF606853-7671-496A-8E4F-DF71F8EC918B}' in style_id:
                            tbl[0][-1].text=new_style_id7
                #１，２、５、９、１０行目を４行目（見やすいやつ）に変更

                for c in range(len(tablee.columns)):
                    cell1 = tablee.cell(0, c)
                    cell1.vertical_anchor = MSO_ANCHOR.MIDDLE
                    pg = cell1.text_frame.paragraphs[0]
                    pg.alignment = PP_ALIGN.CENTER
                for r in range(len(tablee.rows)):
                    cell2 = tablee.cell(r, 0)
                    cell2.vertical_anchor=MSO_ANCHOR.MIDDLE
                    pg = cell2.text_frame.paragraphs[0]
                    pg.alignment = PP_ALIGN.CENTER
                #一番上と一番左のセルの文字を、セルの中心へ
                    
                for x in range(len(tablee.columns)):
                    for y in range(len(tablee.rows)):
                        cell3 = tablee.cell(y, x)
                        pg = cell3.text_frame.paragraphs[0]
                        cell_text=pg.text
                        #print('isdigit:', cell_text.isdigit())#数字かどうかを確かめる用（確認用）
                        if cell_text.isdigit():#もし数字なら
                            cell3.vertical_anchor=MSO_ANCHOR.MIDDLE
                            pg.alignment = PP_ALIGN.RIGHT#数字は右揃え＆上下中央揃え
                        
                        for run in pg.runs:
                            run.font.italic=None
                            run.font.underline=None#イタリックと下線をなくす
                if"清聴" in shape.text or '静聴'in shape.text:
                    xml_slides = prs.slides._sldIdLst
                    slides = list(xml_slides)
                    xml_slides.remove(slides[-1])
                for pg in shape.text_frame.paragraphs:
                    if pg.alignment!=1:
                        pg.alignment = 1
                        lef=1
                    for run in pg.runs:
                        if run.font.italic == True:
                            ita=1
                            run.font.italic=None
                        if run.font.bold == True:
                            bol=1
                            run.font.bold=None
                        if run.font.underline == True:
                            und=1
                            run.font.underline=None
        self.displacement()
        self.replace()
        prs.save(file_22)
        if self.check==0:
            if ita ==1:
                sent=sent+'\nイタリックはやめましょう'
            if bol ==1:
                sent=sent+'\n太字はやめましょう'
            if und==1:
                sent=sent+'\n下線はやめましょう'
            if lef==1:
                sent=sent+'\n左揃えにしましょう'
            if sent!='':
                self.lbl.text=sent
        

    def popup_open(self):
        content = PopupMenu(popup_close=self.popup_close)
        self.popup = Popup(title='選択', content=content, size_hint=(0.75, 0.7), auto_dismiss=False)
        self.popup.open()

    def popup_open2(self):
        content = PopupMenu2(popup_close2=self.popup_close2)
        self.popup = Popup(title='保存', content=content, size_hint=(0.75, 0.7), auto_dismiss=False)
        if file_1!='aaaaa' and file_1!=[]:
            self.popup.open()
        else:
            self.lbl.text='先にファイルを保存してください'

    def popup_close2(self):
        global file_22
        self.popup.dismiss()
        if file2=='':
            return
        file_22='{}/{}.pptx'.format(file_2,file2)
        self.dia(sent)

    def popup_close(self):
        global file_1,file_2,prs,sld1
        self.popup.dismiss()
        if file_1!='aaaaa' and file_1!=[]:
            self.setting()

    def setting(self):
        global prs,sld1,file_1
        prs = Presentation(file_1[0])
        sld1 = prs.slides[-1]
        self.lbl.text='名前を付けて保存してください'


    def pressButton(self): 
        sm.current = 'select'
        sm.transition.direction='left'
        
    def checkbox_1(self):
        self.check=0
    
    def checkbox_2(self):
        self.check=1
        
    def  displacement(self):
        shp_name=[]
        plc_l=[]
        plc_w=[]
        plc_t=[]
        plc_h=[]
        slide=[]
        shap=[]
        center_x=[]
        center_y=[]
        right=[]
        bottom=[]
        rotation=[]
        cnt=-1
        df=pd.DataFrame({})
        for sld in prs.slides:
            shapes=sld.shapes
            cnt+=1
            count=-1
            for shp in sld.shapes:
                count+=1
                s=''
                ll=int(shp.left)/12700
                tt=int(shp.top)/12700
                ww=int(shp.width)/12700
                hh=int(shp.height)/12700
                rotation.append(shp.rotation)
                slide.append(cnt)
                shap.append(count)
                plc_h.append(hh)
                plc_w.append(ww)
                plc_t.append(tt)
                plc_l.append(ll)
                shp_name.append(shp.name)

        df['slide']=slide
        df['shape']=shap
        df['Name']=shp_name
        df['Left']=plc_l
        df['Top']=plc_t
        df['width']=plc_w
        df['Height']=plc_h
        df['Rotation']=rotation
        for i in df.index:
            center_x.append(df.at[i,'Left']+df.at[i,'width']/2)
            center_y.append(df.at[i,'Top']+df.at[i,'Height']/2)
            right.append(df.at[i,'Left']+df.at[i,'width'])
            bottom.append(df.at[i,'Top']+df.at[i,'Height'])
        df['center_x']=center_x
        df['center_y']=center_y
        df['Right']=right
        df['Bottom']=bottom
        df_l=df.sort_values('center_x')
        df_t=df.sort_values('center_y')
        for i in range(len(df_l.index)):
            ss :int =df_t.index[i]
            sss:int=df_t.index[i-1]
            dd:int=df_l.index[i]
            ddd:int=df_l.index[i-1]
            if numpy.abs(df.at[ss,'Left']-df.at[sss,'Right'])<=65 and numpy.abs(df.at[ss,'center_y']-df.at[sss,'center_y'])<=65 and i!=0 and df.at[ss,'slide']==df.at[sss,'slide'] and df.at[ss,'Left']-df.at[sss,'Right']!=0 and df.at[ss,'Top']-df.at[sss,'Bottom']!=0:
                center=df.at[ss,'Top']+df.at[sss,'center_y']-df.at[ss,'center_y']
                df.at[ss,'center_y']=df.at[sss,'center_y']
                df.at[ss,'Top']=center
                prs.slides[df.at[ss,'slide']].shapes[df.at[ss,'shape']].top=Pt(center)
                
            if  numpy.abs(df.at[ss,'center_x']-df.at[sss,'center_x'])<=65 and numpy.abs(df.at[ss,'Top']-df.at[sss,'Bottom'])<=65 and i!=0 and df.at[ss,'slide']==df.at[sss,'slide'] and df.at[ss,'Top']-df.at[sss,'Bottom']!=0 and df.at[ss,'Left']-df.at[sss,'Right']!=0:
                center_2=df.at[ss,'Left']+df.at[sss,'center_x']-df.at[ss,'center_x']
                df.at[ss,'center_x']=df.at[sss,'center_x']
                df.at[ss,'Left']=center_2
                prs.slides[df.at[ss,'slide']].shapes[df.at[ss,'shape']].left=Pt(center_2)
                
            if numpy.abs(df.at[dd,'Left']-df.at[ddd,'Right'])<=65 and numpy.abs(df.at[dd,'center_y']-df.at[ddd,'center_y'])<=65 and i!=0 and df.at[dd,'slide']==df.at[ddd,'slide'] and df.at[dd,'Left']-df.at[ddd,'Right']!=0 and df.at[dd,'Top']-df.at[ddd,'Bottom']!=0:
                center=df.at[dd,'Top']+df.at[ddd,'center_y']-df.at[dd,'center_y']
                df.at[dd,'center_y']=df.at[ddd,'center_y']
                df.at[dd,'Top']=center
                prs.slides[df.at[dd,'slide']].shapes[df.at[dd,'shape']].top=Pt(center)
                
            if  numpy.abs(df.at[dd,'center_x']-df.at[ddd,'center_x'])<=65 and numpy.abs(df.at[dd,'Top']-df.at[ddd,'Bottom'])<=65 and i!=0 and df.at[dd,'slide']==df.at[ddd,'slide'] and df.at[dd,'Top']-df.at[ddd,'Bottom']!=0 and df.at[dd,'Left']-df.at[ddd,'Right']!=0:
                center_2=df.at[dd,'Left']+df.at[ddd,'center_x']-df.at[dd,'center_x']
                df.at[dd,'center_x']=df.at[ddd,'center_x']
                df.at[dd,'Left']=center_2
                prs.slides[df.at[dd,'slide']].shapes[df.at[dd,'shape']].left=Pt(center_2)
                
    def replace(self):
        df=pd.DataFrame({})
        cnt=0
        bold=[]
        ita=[]
        und=[]
        sld=[]
        shp=[]
        para=[]
        fonts=[]
        texts=[]
        runs=[]
        color_1=[]
        color_2=[]
        color_3=[]
        color_t=[]
        color_b=[]


        for sld1 in prs.slides:
            cnt+=1
            cont=0
            for shp1 in sld1.shapes:
                count=0
                cont+=1
                if not shp1.has_text_frame:
                    continue
                for pg1 in shp1.text_frame.paragraphs:
                    count+=1
                    run=0
                    for run1 in pg1.runs:
                        run+=1
                        k=run1.font.size
                        t=run1.text
                        bl=run1.font.bold
                        it=run1.font.italic
                        un=run1.font.underline
                        try:
                            text_color=run1.font.color.rgb
                            color_1.append(int(text_color[0]))
                            color_2.append(int(text_color[1]))
                            color_3.append(int(text_color[2]))
                        except AttributeError:
                            try:
                                text_color_t=run1.font.color.theme_color
                                text_color_bright=run1.font.color.brightness
                                color_1.append(-1)
                                color_2.append(-1)
                                color_3.append(-1)
                            except AttributeError:
                                color_1.append(0)
                                color_2.append(0)
                                color_3.append(0)
                                text_color_t=-1
                                text_color_bright=-1
        
                        if k!=None:
                            fig=int(int(k)/12700)
                        else:
                            fig=None
                        runs.append(run-1)
                        fonts.append(fig)
                        texts.append(t)
                        sld.append(cnt-1)
                        shp.append(cont-1)
                        para.append(count-1)
                        bold.append(bl)
                        ita.append(it)
                        und.append(un)
                        color_t.append(text_color_t)
                        color_b.append(text_color_bright)
        df['slide']=sld
        df['shape']=shp
        df['paragraph']=para
        df['size']=fonts
        df['text']=texts
        df['run']=runs
        df['bold']=bold
        df['italic']=ita
        df['underline']=und
        df['Color_R']=color_1
        df['Color_G']=color_2
        df['Color_B']=color_3
        df['T_Color']=color_t
        df['B_Color']=color_b
        df.fillna({'size':-5},inplace=True)
        for sld1 in prs.slides:
            for shp1 in sld1.shapes:
                if not shp1.has_text_frame:
                    continue
                for pg1 in shp1.text_frame.paragraphs:
                        pg1.text=''
        for index in df.index:
            sld2=prs.slides[df.at[index,'slide']]
            shp2=sld2.shapes[df.at[index,'shape']]
            pg2=shp2.text_frame.paragraphs[df.at[index,'paragraph']]
            new_text=pg2.add_run()
            new_text.text=df.at[index,'text']
            new_text.font.name='MairyoUI'
            new_text.font.bold=df.at[index,'bold']
            new_text.font.italic=df.at[index,'italic']
            new_text.font.underline=df.at[index,'underline']
            if df.at[index,'T_Color']==-1:
                new_text.font.color.rgb=RGBColor(int(df.at[index,'Color_R']),int(df.at[index,'Color_G']),int(df.at[index,'Color_B']))
            else:
                new_text.font.color.theme_color=df.at[index,'T_Color']
                new_text.font.color.brightness=df.at[index,'B_Color']
            if df.at[index,'size']==-5:
                continue
            new_text.font.size=Pt(df.at[index,'size'])

class rrr (Screen):
    def __init__(self, **kwargs):
        super(rrr, self).__init__(**kwargs)
        pass
    def back(self):
        sm.transition.direction='right'
        sm.current = 'menu'

class PopupMenu(Screen):
    popup_close = ObjectProperty(None)
    def filea(self):
        global file_1,file_2,prs,sld1
        file_1=self.vv.selection
        if file_1==[] :
            self.bt.text='先にファイルを選択して'
        else:
            self.popup_close()

class PopupMenu2(Screen):
    popup_close2 = ObjectProperty(None)
    def fileb(self):
        global file_1,file_2,prs,sld1,file2
        file_2=self.vv.path    
        file2=self.input.text
        if file2=='':
            self.input.text='ファイル名を入力してください'
        else:
            self.popup_close2()


class StudyApp(App):
    def build(self):
        sm.add_widget(diagnosis(name='menu'))
        sm.add_widget(rrr(name='select'))
        Window.bind(on_dropfile=self._on_file_drop)
        return sm

    def _on_file_drop(self,window,file_path):
        global file_1
        file_1=file_path.decode('utf-8')
        file_1=file_1.rsplit()
        '\\'.join(file_1)
        diagnosis().setting()
        


if __name__ == '__main__':
    StudyApp().run()
