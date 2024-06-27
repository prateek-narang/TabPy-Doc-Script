import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import sys
import os.path
import base64

# ignoring default='warn'
pd.options.mode.chained_assignment = None

class TableauDocument():
    
    def __init__(self, filePath):
        self.filePath = os.path.normpath(filePath)
        self.xmlRoot = self._get_xml_root()
        self.parameters = self._get_parameters()
        self.calculations = self._get_calculations()

    def _get_xml_root(self):
        '''To convert twbx/twb file to XML root element'''
        filePath = self.filePath
        if filePath[-4:]=='.twb':
            root = ET.parse(filePath).getroot()
        elif filePath[-4:]=='twbx':
            with open(filePath, "rb") as binfile:
                twbx = zipfile.ZipFile(binfile)
                name = [w for w in twbx.namelist() if w.find(".twb") != -1][0]
                unzip = twbx.open(name)
                xml = ET.parse(unzip)
                xml.write(name[:-4] + ".xml")   # Writing corresponding xml file
                root = xml.getroot()
        else:
            print('Note: The file should be .twbx OR a .twb')
            sys.exit()
        return root

    def _get_parameters(self):
        '''To get parameters in a dataframe from the root element'''
        temp = []
        parameters = self.xmlRoot.findall("./datasources/datasource[@name='Parameters']/column")
        for element in parameters:        
            temp.append({
                'caption': element.get('caption'),
                'datatype': element.get('datatype'),
                'name': element.get('name'),
                'default_value': element.get('value'),
                'param-domain-type': element.get('param-domain-type'),
                'role': element.get('role'),
                'type': element.get('type'),
                'members': self.find_members(element)
            })
        df = pd.DataFrame(temp)
        return df

    def _get_calculations(self):
        '''To get calculations in a dataframe from the root element'''
        temp = []
        datasources = self.xmlRoot.findall("./datasources/datasource[@name!='Parameters']")
        
        # get all calculations for each data source
        for datasource in datasources:
            calc = datasource.findall("./*calculation[@class='tableau']/..[@name]")
            for element in calc:
                temp.append({
                    'datasource': self.extract_alias_name(datasource),
                    'caption': self.extract_alias_name(element),
                    'name': element.get('name'),
                    'role': element.get('role'),
                    'datatype': element.get('datatype'),
                    'type': element.get('type'),
                    'formula': element.find('./calculation').get('formula')
                })
        df = pd.DataFrame(temp)
        if len(df)==0:  # if there are no calculations in dashboard
            return df
        
        # update calculation formula & add impacted calculations
        temp_df = pd.DataFrame()
        for ds in df['datasource'].unique():
            calcs_sub = df[df['datasource']==ds]
            calcs_sub = self.update_calculation_formula(calcs_sub)
            calcs_sub = self.add_impacted_fields(calcs_sub)
            temp_df = pd.concat([temp_df, calcs_sub])
        
        df = temp_df.reset_index(drop=True)
        return df

    def find_members(self, parameter):
        member = []
        members = parameter.findall("./members/member")
        for item in members:
            member.append(item.get('value'))
        return ",".join(member)
        
    def update_calculation_formula(self, calcs_sub):
        '''To update the formula field, to replace tableau identifiers with their user defined names'''
        search_dict = self.create_identifier_dict(calcs_sub)
        search_dict.update(self.create_identifier_dict(self.parameters))
        
        for i in range(len(calcs_sub)):
            calcs_sub = calcs_sub.reset_index(drop=True)
            cell = calcs_sub['formula'][i]
            for key in list(search_dict.keys()):
                if key in cell:
                    cell = cell.replace(key, "[" + search_dict[key] + "]")
            calcs_sub['formula'][i] = cell  
        return calcs_sub

    def add_impacted_fields(self, calcs_sub):
        '''To add a list of corresponding impacted calculations for each calculation'''
            
        calcs_sub['impacts'] = pd.Series()
        for i in range(len(calcs_sub)):
            calcs_sub['impacts'][i] = ""
            _caption = "[" + calcs_sub['caption'][i] + "]"
            for j in range(len(calcs_sub)): 
                if _caption in calcs_sub['formula'][j]:
                    calcs_sub['impacts'][i] += (",[" + calcs_sub['caption'][j] + "]")
            
            calcs_sub['impacts'][i] = calcs_sub['impacts'][i][1:]

        return calcs_sub

    def output_to_excel(self):
        '''To write to multiple sheets in excel using ExcelWriter'''
        with pd.ExcelWriter("./Results.xlsx") as writer:
            self.parameters.to_excel(writer, sheet_name="Parameters", index=False)
            self.calculations.to_excel(writer, sheet_name="Calculations", index=False)

    def create_identifier_dict(self, df):
        df = df[['name','caption']]
        df.index = df['name']
        df.drop(columns='name', inplace=True)
        dictn =  df.to_dict()['caption']
        return dictn
        
    def extract_alias_name(self, element):
        if element.get('caption') is None:
            element = element.get('name')
            val = element.replace("[","").replace("]","")
            return val
        else:
            return element.get('caption')

    def generate_thumbnails(self):
        tmb = self.xmlRoot.findall('./thumbnails/thumbnail')
        if tmb==[]:
            print("No thumbnail images found")
        else:
            for item in tmb:
                name = item.get("name")
                with open("./img/" + name + ".png", "wb") as fh:
                    fh.write(base64.decodebytes(item.text.encode()))