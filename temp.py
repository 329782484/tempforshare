import os
import time
import uuid
import zipfile
from io import BytesIO
from tkinter.filedialog import askopenfilenames
import re

from lxml import etree
import fastexcel
# from multiprocessing import Pool
from joblib import Parallel, delayed

def uuidstr():
    return str(uuid.uuid1())

# from memory_profiler import profile

# @profile(precision=4)
class S7Tia():

    def __init__(self, filepath):
        self.file = filepath
        self.fname = ''
        # self.filepath = ''
        self.points={}
        self.points['TIA_BlockInterfaceErrorList'] = {}
        self.points['TIA_SEDBtitle']={}
        self.tagsused = {}
        self.tagslot={}

        self.fileinfo()

        self.tagused()
        self.findtagslot()

        # func_list=['self.tagused','self.findtagslot']
        # results = Parallel(n_jobs=-1, backend='threading')(delayed(self.func)(fn) for fn in func_list)

        self.device()
        self.block()
        self.blockitem()
        self.blockinterface()
        self.tag()

        func_list=['self.device','self.block','self.blockitem','self.blockinterface','self.tag']
        results = Parallel(n_jobs=-1, backend='threading')(delayed(self.func)(fn) for fn in func_list)

    def func(self,fn):
        eval(fn)()

    def fileinfo(self):
        filepath, fullflname = os.path.split(self.file)
        fname, _ = os.path.splitext(fullflname)
        self.fname = fname
        # self.filepath = filepath
        # ext = ext

    def FindTypeIdent(self,tnd):
        tndlst=[]
        if tnd != None:
            TInd=tnd.xpath('./Attribute[@Name="TypeIdentifier"][text()!=" "]')
            if len(TInd)>0:
                tndlst.append(''.join(TInd[0].xpath('text()')))
            else:
                tndname=tnd.xpath('@Name')
                if len(tndname)>0:
                    tndlst.append(tndname[0])
                else:
                    tndlst.append(tnd.tag)
                tndlst.extend(self.FindTypeIdent(tnd.getparent()))
        return tndlst

    def findtagslot(self):
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if f == 'Hardware.xml':
                root = etree.parse(BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                for address in root.xpath('.//Address/Attribute[@Name="StartAddress"]/parent::*'):
                    startaddress=''.join(address.xpath('./Attribute[@Name="StartAddress"]/text()'))
                    iotype=''
                    positionnumber=''
                    devicename=''
                    devicetype=''
                    length=''
                    if startaddress!='-1':
                        length=''.join(address.xpath('./Attribute[@Name="Length"]/text()'))
                        iotype=''.join(address.xpath('./Attribute[@Name="IoType"]/text()'))
                        positionnumber=''.join(address.getparent().getparent().xpath('Attribute[@Name="PositionNumber"]/text()'))
                        devicetype=self.dttxt( ''.join(address.getparent().getparent().xpath('Attribute[@Name="TypeIdentifier"]/text()')))
                        # for racknode in address.xpath('./ancestor::*/Attribute[@Name="TypeName"][text()="Rack" or text()="Rail"]/parent::*'):
                        for racknode in address.xpath('./ancestor::*/Attribute[@Name="TypeName"][text()="Rack" or text()="Rail"]/parent::*') + address.xpath('./ancestor::*/Attribute[@Name="TypeIdentifier"][text()="System:Rack.KP32F"]/parent::*'):
                            for devicenode in racknode.xpath('./*'):
                                notes=devicenode.xpath('.//Node/Attribute[@Name="PnDeviceName"]/parent::*')
                                if len(notes)>0:
                                    for Note in devicenode.xpath('.//Node/Attribute[@Name="PnDeviceName"]/parent::*'):
                                        devicename = ''.join(devicenode.xpath('@Name'))

                        if iotype.lower()=='input':
                            tp='I'
                        elif iotype.lower()=='output':
                            tp='Q'
                        else:
                            tp='?'

                        startaddress=int(startaddress)*8
                        for ad in range(startaddress,startaddress+int(length)):
                            devicename=devicename.upper()
                            result=re.search("\d\d\d[A-Z][A-Z][A-Z-]\d\d\d",devicename)
                            if result is not None:
                                dn=result.group()
                                dnn=dn.replace('-','_')
                                devicename=devicename.replace(dn,dnn)
                            self.tagslot[tp+str(ad)]=['NET001='+devicename+'.'+str(positionnumber),devicetype]

    def device(self):
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if f == 'Hardware.xml':
                # function 2
                root = etree.parse(
                    BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                for racknode in root.xpath('//Attribute[@Name="TypeName"][text()="Rack" or text()="Rail"]/parent::*'):
                    for devicenode in racknode.xpath('./*'):
                        notes=devicenode.xpath('.//Node/Attribute[@Name="PnDeviceName"]/parent::*')
                        if len(notes)>0:
                            for Note in devicenode.xpath('.//Node/Attribute[@Name="PnDeviceName"]/parent::*'):
                                Item = {}
                                Item['TIA_PLC'] = self.fname
                                Item['ProjectDeviceName'] = ''.join(devicenode.xpath('@Name'))
                                Item['PnDeviceName'] = ''.join(Note.xpath(
                                    './Attribute[@Name="PnDeviceName"]/text()')[0].split('.')[0])
                                Item['Address'] = ''.join(Note.xpath(
                                    './Attribute[@Name="Address"]/text()'))
                                ###################################new#######
                                Item['TypeName'] = ''.join(devicenode.xpath('./Attribute[@Name="TypeName"]/text()'))
                                ###################################new#######

                                if 'TIA_hardware' in self.points.keys():
                                    self.points['TIA_hardware'][uuidstr()] = Item
                                else:
                                    self.points['TIA_hardware'] = {}
                                    self.points['TIA_hardware'][uuidstr()] = Item

                ######## new code #####
                notelist = []
                for Note in root.xpath('//Attribute'):
                    notelist.extend(Note.xpath('./parent::*'))
                notelist = list(set(notelist))

                for nl in notelist:
                    itemname = ''
                    Item = {}
                    Item['TIA_PLC'] = self.fname
                    for hdnote in nl.xpath('./ancestor-or-self::*'):
                        attexist=hdnote.xpath('@Name')
                        if len(attexist)>0:
                            itemname=itemname+'.'+attexist[0]
                        else:
                            itemname=itemname+'.'+hdnote.tag
                    Item['notepath'] = itemname
                    tistr=self.FindTypeIdent(nl)
                    tistr.reverse()
                    # tistr=list(tistr)
                    tistr='.'.join(tistr)
                    Item['notetypeident'] = tistr
                    for Att in nl.xpath('./Attribute'):
                        Item[Att.xpath('@Name')[0]] = ''.join(
                            Att.xpath('text()'))

                    if 'TIA_hdw' in self.points.keys():
                        self.points['TIA_hdw'][uuidstr()] = Item
                    else:
                        self.points['TIA_hdw'] = {}
                        self.points['TIA_hdw'][uuidstr()] = Item

    def block(self):
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if ('Blocks/' in f) and f[-3:].upper() == 'XML':
                # function 2
                root = etree.parse(
                    BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                typ = '(//SW.Blocks.FC|//SW.Blocks.FB|//SW.Blocks.OB|//SW.Blocks.GlobalDB|//SW.Blocks.InstanceDB)'
                Item = {}
                Item['TIA_PLC'] = self.fname
                Item['Name'] = ''
                for Att in root.xpath(typ+'/AttributeList/*'):
                    Item[Att.tag] = ''.join(Att.xpath('text()'))

                results = root.xpath(
                    '/Document/*/ObjectList/MultilingualText[@CompositionName="Title"]/ObjectList/MultilingualTextItem/AttributeList')
                for result in results:
                    Item['Title_'+result.xpath('Culture')
                         [0].text] = result.xpath('Text')[0].text

                results = root.xpath(
                    '/Document/*/ObjectList/MultilingualText[@CompositionName="Comment"]/ObjectList/MultilingualTextItem/AttributeList')
                for result in results:
                    Item['Comment_'+result.xpath('Culture')
                         [0].text] = result.xpath('Text')[0].text

                if 'TIA_blocks' in self.points.keys():
                    self.points['TIA_blocks'][uuidstr()] = Item
                else:
                    self.points['TIA_blocks'] = {}
                    self.points['TIA_blocks'][uuidstr()] = Item


    def blockitem(self):
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if ('Blocks/' in f) and f[-3:].upper() == 'XML':
                # function 2
                root = etree.parse(
                    BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                GlobalDBName = ''.join(root.xpath(
                    '//SW.Blocks.GlobalDB/AttributeList/Name/text()'))
                for member in root.xpath('//SW.Blocks.GlobalDB/AttributeList/Interface//Member'):
                    Item = {}
                    Item['TIA_PLC'] = self.fname
                    itemname = ''
                    for m in member.xpath('./ancestor-or-self::Member/@Name'):
                        itemname = itemname+'.'+m
                    Item['Name'] = GlobalDBName+itemname
                    Item['Datatype'] = ''.join(member.xpath('@Datatype'))
                    Item['used'] = 'Yes' if Item['Name'] in self.tagsused else 'No'
                    for comm in member.xpath('./Comment/*'):
                        Item[comm.attrib['Lang']] = comm.text
                    Item['Remanence'] = ''.join(member.xpath('@Remanence'))
                    Item['Accessibility'] = ''.join(
                        member.xpath('@Accessibility'))
                    for att in member.xpath('./AttributeList/*'):
                        Item[att.attrib['Name']] = att.text

                    if 'TIA_blocksitems' in self.points.keys():
                        self.points['TIA_blocksitems'][uuidstr()] = Item
                    else:
                        self.points['TIA_blocksitems'] = {}
                        self.points['TIA_blocksitems'][uuidstr()] = Item

    def Ic(self, fn, wire,blockuid):
        icid=''
        if len(wire.xpath('./IdentCon'))==0 and len(wire.xpath('./NameCon'))>1:
            NameConuid=wire.xpath('./NameCon[@UId!="'+blockuid+'"]')[0].xpath('@UId')[0]
            try:
                wire2 = fn.xpath('./Wires/Wire/NameCon[@UId="'+NameConuid+'"][@Name="operand" or @Name="in1"]/..')[0]
                icid = self.Ic(fn, wire2,NameConuid)
            except:
                icid = 'notfind_'+NameConuid
        else:
            try:
                icid = wire.xpath('./IdentCon')[0].xpath('@UId')[0]
            except:
                pass
        return(icid)

        # if wire.xpath('./*')[0].tag == 'NameCon' and wire.xpath('./*')[1].tag == 'NameCon':
        #     wireid = wire.xpath('./*')[0].xpath('@UId')[0]
        #     # wireid = wire.xpath('./NameCon')[0].xpath('@UId')[0]
        #     # 2020-02-18 update
        #     try:
        #         wire2 = fn.xpath(
        #             './Wires/Wire/NameCon[@UId="'+wireid+'"][@Name="operand" or @Name="in1"]/..')[0]
        #         id = self.Ic(fn, wire2)
        #     except:
        #         id = 'notfind_'+wireid
        # else:
        #     try:
        #         id = wire.xpath('./*')[0].xpath('@UId')[0]
        #     except:
        #         pass

        # return(id)

    def typematch(self,sedbname):
        tp=''
        if re.match("^SEDB_\d\d\dBR_\d\d\d-FN1$",sedbname) is not None:
            # SEDB_001BR_010-FN1
            tp='VI- EBF'
        elif re.match("^SEDB_\d\d\dSTU\d\d\d$",sedbname) is not None:
            # SEDB_040STU002
            tp='estop and door'
        elif re.match("^SEDB_\d\d\dPNE\d\d\d$",sedbname) is not None:
            # SEDB_001PNE001
            tp='compressed air'
        elif re.match("^SEDB_\d\d\dRB_\d\d\d-US2$",sedbname) is not None:
            # SEDB_040RB_200-US2
            tp='F-PME US2'
        elif re.match("^SEDB_\d\d\dRB_\d\d\d-Safe$",sedbname) is not None:
            # SEDB_050RB_100-Safe
            tp='Safezone 1'
        elif re.match("^SEDB_\d\d\dRB_\d\d\d$",sedbname) is not None:
            # SEDB_001BR_010-FN1
            tp='estop and robot'
        elif re.match("^SEDB_\d\d\d[A-Z][A-Z][A-Z_]\d\d\d-MA1$",sedbname) is not None:
            # SEDB_190SGM002-MA1
            tp='Release Lenze'
        elif re.match("^SEDB_\d\d\dSGM\d\d\d-FN1$",sedbname) is not None:
            # SEDB_001BR_010-FN1
            tp='Safety window'
        elif re.match("^SEDB_\d\d\dNH_\d\d\d-FN1$",sedbname) is not None:
            # SEDB_001BR_010-FN1
            tp='estop'
        elif re.match("^SEDB_\d\d\dBR_\d\d\d-FN2$",sedbname) is not None:
            # SEDB_001BR_010-FN1
            tp='estop'
        # elif re.match("^SEDB_\d\d\dBR_\d\d\d-FN1$",sedbname) is not None:
        #     # SEDB_001BR_010-FN1
        #     tp='VI- EBF'
        # elif re.match("^SEDB_\d\d\dBR_\d\d\d-FN1$",sedbname) is not None:
        #     # SEDB_001BR_010-FN1
        #     tp='VI- EBF'
        # elif re.match("^SEDB_\d\d\dBR_\d\d\d-FN1$",sedbname) is not None:
        #     # SEDB_001BR_010-FN1
        #     tp='VI- EBF'
        # elif re.match("^SEDB_\d\d\dBR_\d\d\d-FN1$",sedbname) is not None:
        #     # SEDB_001BR_010-FN1
        #     tp='VI- EBF'
        # elif re.match("^SEDB_\d\d\dBR_\d\d\d-FN1$",sedbname) is not None:
        #     # SEDB_001BR_010-FN1
        #     tp='VI- EBF'

        return tp


    def dttxt(self,gsd):
        dt=''
        if 'MURRELEKTRONIK' in gsd:
            dt='MVK_MPNIO'
        elif 'FESTO' in gsd:
            dt='Festo'
        elif '6PA00' in gsd:
            dt='F-PME'
        return dt

    def sedbtitle(self,plcname,blockname,sedbname,parameter,ad):
        if 'SEDB_' not in sedbname:
            return
        if ad=='':
            title=''
        else:
            title=self.tagslot[ad][0]+', '+self.typematch(sedbname)+self.tagslot[ad][1]

        if blockname in ['iFG_F_2-HAND','iFG_F_ESTOP_8','iFG_F_ESTOP_X','iFG_F_ESTOP_xBYPASS','iFG_F_MODE_3POS','iFG_F_VI_3POS','iFG_F_VI_DESK_3POS']:
            if parameter!='in_1':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_BASIC_DIAG"]:
            if parameter!='out_1':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_BASIC_LENZE",'iFG_F_EXTENDED_LENZE']:
            if parameter!='enableSTO':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_CENTRAL_LOCKING"]:
            if parameter!='unlock':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_EUCHNER_GATE"]:
            if parameter!='inEStopAck':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_FEEDBACK_2C"]:
            if parameter!='channel':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_FEEDBACK_LOOP"]:
            if parameter!='C_1':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_INLAID_AREA_L_CURTAIN"]:
            if parameter!='in_1_lightCurtain':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_INLAID_AREA_PROT_GATE"]:
            if parameter!='in_1_protectionGate':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_PALLET_TRANSFER"]:
            if parameter!='bws1':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_POSITION_xBYPASS"]:
            if parameter!='position_1':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_ROBOT_SAFE_BASIC"]:
            if parameter!='outIfcSafe':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_ROBOT_SAFE_OP"]:
            if parameter!='enableFlex_1':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_ROBOT_TECHNO_PS"]:
            if parameter!='enablePS':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_S3000"]:
            if parameter!='protectField':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_SAFETY_PINS_EN"]:
            if parameter!='locked':
            #     title=self.tagslot[ad]
            # else:
                return
        elif blockname in ["iFG_F_TOOL_STOWAGE"]:
            if parameter!='stowageRoom_1':
            #     title=self.tagslot[ad]
            # else:
                return

        P = {}
        P['TIA_PLC'] = plcname
        P['Block_Name'] = blockname
        P['Block_Instance'] = sedbname
        P['title'] = title
        self.points['TIA_SEDBtitle'][uuidstr()] = P

    def blockinterface(self):
        # exceplist=['Contact', 'O', 'Coil', 'TrCoil', 'SvCoil', 'IlCoil', 'Move', 'Rs', 'RCoil', 'Sr', 'SCoil', 'TON', 'PBox', 'Eq', 'Not', 'NBox', 'Gt', 'Ge', 'Ne', 'GetInstanceName', 'Jump', 'FILL', 'Return', 'Add', 'TOF', 'Le', 'Convert','TP', 'MIN', 'Strg_TO_Chars','FillBlockI','Lt', 'S_Move', 'Sub', 'Div', 'Mul', 'GETIO_PART', 'SETIO_PART', 'PContact', 'GETIO', 'SETIO', 'T_CONV', 'JumpNot','Deserialize', 'Serialize', 'Runtime', 'NContact', 'R_TRIG', 'RDREC', 'LOG2GEO', 'Swap','InRange', 'CTU', 'Calc', 'PCoil']
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if f[-3:].upper() == 'XML':
                # function 2
                root = etree.parse(
                    BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                FlgNet = root.xpath('//Call/ancestor::FlgNet')
                for fn in FlgNet:
                    for Call in fn.xpath('./Parts/Call'):
                        Calluid = Call.xpath('@UId')[0]
                        CallName = Call.xpath('./CallInfo/@Name')[0]
                        Instance = '.'.join(Call.xpath('./CallInfo/Instance/Component/@Name'))
                        Instance_Type= '.'.join(Call.xpath('./CallInfo/Instance/@Scope'))
                        P = {}
                        P['TIA_PLC'] = self.fname
                        P['Block_Name'] = CallName
                        P['Block_Instance'] = Instance
                        P['Instance_Type'] = Instance_Type

                        for wire in fn.xpath('./Wires/Wire[not(Powerrail)]/NameCon[@UId="'+Calluid+'"]/..'):
                            Parameter=wire.xpath('./NameCon[@UId="'+Calluid+'"]/@Name')[0]
                            IdentConuid = self.Ic(fn, wire,Calluid)
                            Component = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Component/@Name'))
                            ComponentadI = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Address[@Area="Input"]/@BitOffset'))
                            ComponentadQ = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Address[@Area="Output"]/@BitOffset'))
                            if ComponentadI!='':
                                ComponentadI='I'+ComponentadI
                            if ComponentadQ!='':
                                ComponentadQ='Q'+ComponentadQ
                            Componentad=ComponentadI+ComponentadQ
                            ConstantValue = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//ConstantValue/text()'))
                            Constant = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Constant/@Name'))
                            P[Parameter] = Component+ConstantValue+Constant

                            self.sedbtitle(P['TIA_PLC'] ,P['Block_Name'],P['Block_Instance'],Parameter,Componentad)


                        if P['Block_Name'] in self.points.keys():
                            self.points[P['Block_Name']
                                                 ][uuidstr()] = P
                            ##check fault
                            if "FGDB_" in P['Block_Instance']:
                                fgname=P['Block_Instance'].replace('FGDB_','')
                                for k in P.keys():
                                    result=re.match("\d\d\d[A-Z][A-Z][A-Z_]\d\d\d",P[k])
                                    if result is not None:
                                        if result.group()[3:6] not in ['XSR','PNE']:
                                            if fgname not in P[k]:
                                                self.points['TIA_BlockInterfaceErrorList'][uuidstr()] = P
                                                break
                        else:
                            self.points[P['Block_Name']] = {}
                            self.points[P['Block_Name']
                                                 ][uuidstr()] = P
                            ##check fault
                            if "FGDB_" in P['Block_Instance']:
                                fgname=P['Block_Instance'].replace('FGDB_','')
                                for k in P.keys():
                                    result=re.match("\d\d\d[A-Z][A-Z][A-Z_]\d\d\d",P[k])
                                    if result is not None:
                                        if result.group()[3:6] not in ['XSR','PNE']:
                                            if fgname not in P[k]:
                                                self.points['TIA_BlockInterfaceErrorList'][uuidstr()] = P
                                                break

                FlgNet = root.xpath('//Part/ancestor::FlgNet')
                for fn in FlgNet:
                    for Part in fn.xpath('./Parts/Part[Instance]'):
                        Partuid = Part.xpath('@UId')[0]
                        PartName = Part.xpath('./@Name')[0]
                        # new for exception 2020-08-23
                        # if PartName in exceplist:
                        #     continue
                        Instance = '.'.join(Part.xpath('./Instance/Component/@Name'))
                        Instance_Type= '.'.join(Part.xpath('./CallInfo/Instance/@Scope'))
                        P = {}
                        P['TIA_PLC'] = self.fname
                        P['Block_Name'] = PartName
                        P['Block_Instance'] = Instance
                        P['Instance_Type'] = Instance_Type

                        for wire in fn.xpath('./Wires/Wire[not(Powerrail)]/NameCon[@UId="'+Partuid+'"]/..'):
                            Parameter=wire.xpath('./NameCon[@UId="'+Partuid+'"]/@Name')[0]
                            IdentConuid = self.Ic(fn, wire,Partuid)
                            Component = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Component/@Name'))
                            ComponentadI = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Address[@Area="Input"]/@BitOffset'))
                            ComponentadQ = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Address[@Area="Output"]/@BitOffset'))
                            if ComponentadI!='':
                                ComponentadI='I'+ComponentadI
                            if ComponentadQ!='':
                                ComponentadQ='Q'+ComponentadQ
                            Componentad=ComponentadI+ComponentadQ
                            ConstantValue = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//ConstantValue/text()'))
                            Constant = '.'.join(
                                fn.xpath('./Parts/Access[@UId="'+IdentConuid+'"]//Constant/@Name'))
                            P[Parameter] = Component+ConstantValue+Constant

                            self.sedbtitle(P['TIA_PLC'] ,P['Block_Name'],P['Block_Instance'],Parameter,Componentad)

                        if P['Block_Name'] in self.points.keys():
                            self.points[P['Block_Name']
                                                 ][uuidstr()] = P
                            ##check fault
                            if "FGDB_" in P['Block_Instance']:
                                fgname=P['Block_Instance'].replace('FGDB_','')
                                for k in P.keys():
                                    result=re.match("\d\d\d[A-Z][A-Z][A-Z_]\d\d\d",P[k])
                                    if result is not None:
                                        if result.group()[3:6] not in ['XSR','PNE']:
                                            if fgname not in P[k]:
                                                self.points['TIA_BlockInterfaceErrorList'][uuidstr()] = P
                                                break
                        else:
                            self.points[P['Block_Name']] = {}
                            self.points[P['Block_Name']
                                                 ][uuidstr()] = P
                            ##check fault
                            if "FGDB_" in P['Block_Instance']:
                                fgname=P['Block_Instance'].replace('FGDB_','')
                                for k in P.keys():
                                    result=re.match("\d\d\d[A-Z][A-Z][A-Z_]\d\d\d",P[k])
                                    if result is not None:
                                        if result.group()[3:6] not in ['XSR','PNE']:
                                            if fgname not in P[k]:
                                                self.points['TIA_BlockInterfaceErrorList'][uuidstr()] = P
                                                break

    def tag(self):
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if ('TagTables/' in f) and f[-3:].upper() == 'XML':
                # function 2
                root = etree.parse(
                    BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                PlcTagTable = root.xpath('//SW.Tags.PlcTagTable')
                for PTT in PlcTagTable:
                    PlcTagTableName = PTT.xpath(
                        './AttributeList/Name/text()')[0]
                    for PT in PTT.xpath('./ObjectList/SW.Tags.PlcTag'):
                        Item = {}
                        Item['TIA_PLC'] = self.fname
                        Item['Name'] = ''
                        Item['used'] = 'No'
                        Item['PlcTagTableName'] = PlcTagTableName
                        for Att in PT.xpath('./AttributeList/*'):
                            Item[Att.tag] = ''.join(Att.xpath('text()'))

                        results = PT.xpath(
                            './ObjectList/MultilingualText[@CompositionName="Comment"]/ObjectList/MultilingualTextItem/AttributeList')
                        for result in results:
                            Item['Comment_'+result.xpath('Culture')
                                 [0].text] = result.xpath('Text')[0].text

                        Item['used'] = 'Yes' if Item['Name'] in self.tagsused else 'No'
                        # self.tags[uuidstr()] = Item

                        if 'TIA_tags' in self.points.keys():
                            self.points['TIA_tags'][uuidstr()] = Item
                        else:
                            self.points['TIA_tags'] = {}
                            self.points['TIA_tags'][uuidstr()] = Item

    def tagused(self):
        Zip = zipfile.ZipFile(self.file)
        for f in Zip.namelist():
            if f[-3:].upper() == 'XML':
                # function 2
                root = etree.parse(
                    BytesIO(Zip.open(f).read().replace(b'xmlns', b'remove_ns')))

                syms = root.xpath('//FlgNet/Parts/Access/Symbol')
                for sym in syms:
                    Item = {}
                    Item['TIA_PLC'] = self.fname
                    Item['Name'] = '.'.join(sym.xpath('./Component/@Name'))
                    self.tagsused[Item['Name']] = Item

def Tia2excel(td, path):

    wb = fastexcel.Workbook()
    for tk, tv in td.items():
        keystmp = []
        for key in tv.keys():
            keystmp += tv[key].keys()
        keys = list(set(keystmp))
        keys.sort(key=keystmp.index)

        # # only integra standard
        if tk[0:1] == 'i' or True:
            ws = wb.addsheet(tk)
            ws.addrow(['dic_uuid']+keys)
            for item in tv.keys():
                at = [item]
                for key in keys:
                    at.append(tv[item].get(key, ''))
                ws.addrow(at)

        if tk == 'iFC_SYS_EXTERNAL_ALARM':
            ws = wb.addsheet(tk+'_check')
            ws.addrow(['TIA_PLC', 'fct', 'ack', 'triggerMessage', 'message'])
            for item in tv.keys():
                TIA_PLC = tv[item].get('TIA_PLC', '')
                fct = tv[item].get('fct', '')
                ack = tv[item].get('ack', '')
                for i in range(1, 9):
                    at = [TIA_PLC, fct, ack]
                    at.append(tv[item].get('triggerMessage_'+str(i), ''))
                    at.append(tv[item].get('message_'+str(i), ''))
                    ws.addrow(at)

    wb.save(path,fct=2)

if __name__ == "__main__":
    filenames = askopenfilenames(
        title="Select Files", filetypes=[("Files", "*.S7Tia")])
    files = list(filenames)

    lowerdic = {}
    points={}
    openpath = ''

    for file in files:
        S7 = S7Tia(file)
        openpath, _ = os.path.split(file)
        # openpath = S7.filepath
        for key, value in S7.points.items():
            # bug fix (mix upper and lower)
            if key.lower() in lowerdic.keys():
                points[lowerdic[key.lower()]] = {
                    **points[lowerdic[key.lower()]], **value}
            else:
                points[key] = {}
                points[key] = {**points[key], **value}
                lowerdic[key.lower()] = key

    if len(files) > 0:
        print('output excel')
        timestr = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
        Tia2excel(points, openpath+r"\\"+'TIA_points_'+timestr+'.xlsx')

        print('Finish!')

        os.startfile(openpath)


# 分析程序：
# python -m cProfile -o result.Prof Tia_read.py
# snakeviz result.Prof

# @profile
# kernprof -l -v Tia_read.py

# @profile
# python -m memory_profiler Tia_read.py
