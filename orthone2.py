import pandas as pd
import re
import numpy as np
import PyPDF2
import os
import re

class make_excel ():

    def extract_data_from_pdf(self, file_path):
        data = []
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                data.extend(page.extract_text().split())

        return data

    def newdata(self,ex_data):

        self.name = []
        self.name.extend([ex_data[0], ex_data[1], ex_data[4]])

        self.manuf = []

        self.cont = []
        self.cont.append(ex_data[3])

        self.obs = []
        self.obs.append(ex_data[38])
        self.time =[]
        print(self.obs)
        self.totaldata = []
        self.distance =[]
        self.rpd = []
        self.dap = []
        self.fm =[]
        self.pet = []
        self.drp = []
        self.ppa = []
        self.psa = []
        self.xfm = []
        self.cfa = []
        self.xrftm = []
        self.pr = []
        self.kvp = []
        self.xrtc = []
        n = 0
        self.ext = []
        self.exp= []
        self.pulw = []
        self.ird = []
        self.dstd = []
        self.cfh = []
        self.cfw = []
        print(ex_data)
        for i in range(0, len(ex_data)):
            if re.match(r'Manufacturer', ex_data[i]):
                self.manuf.append(ex_data[i+2])
            elif re.match(r'Total:+\d|Total:\w',ex_data[i]):
                self.totaldata.append(ex_data[i]) #Dose Area Product Total, Dose (RP) Total,
                # Fluoro Dose Area Product Total, Fluoro Dose (RP) Total, Acquisition Dose Area Product Total
            elif re.match(r'Point:+\d',ex_data[i]):
                self.distance.append(ex_data[i]) #Distance Source to Reference Point
            elif re.match(r'Time:+\d',ex_data[i]):
                self.time.append(ex_data[i]) # Total Fluoro Time (s), Total Acquisition Time,
                # Exposure time
            elif re.match(r'Definition:+\d|Deﬁnition:+\d',ex_data[i]):
                self.rpd.append(ex_data[i]) #Reference Point Definition (cm)
            elif re.match(r"Product:+\d",ex_data[i]):
                self.dap.append(ex_data[i]) #Dose area product
            elif re.match(r"\(RP\):\d",ex_data[i]):
                self.drp.append(ex_data[i]) # Dose (RP) (Gy)
            elif re.match(r"Angle:-?\d+(\.\d+)?",ex_data[i]) and "Primary" == ex_data[i-1]:
                self.ppa.append(ex_data[i]) # Positioner Primary Angle (deg)
            elif re.match(r"Angle:-?\d+(\.\d+)?",ex_data[i]) and "Secondary" == ex_data[i-1]:
                self.psa.append(ex_data[i])#Positioner Secondary Angle (deg)
            elif re.match(r"Material:\w",ex_data[i]):
                self.xfm.append(ex_data[i]) #X-Ray Filter Material
            elif re.match(r"Mode:\w",ex_data[i]):
                self.fm.append(ex_data[i]) # Fluoro Mode
            elif re.match(r"Area:\d",ex_data[i]) and "Field" == ex_data[i-1]:
                self.cfa.append(ex_data[i]) # Collimated Field Area (m2)
            elif re.match(r"Rate:\d",ex_data[i]):
                self.pr.append(ex_data[i]) #Pulse Rate (pulse/s)
            elif re.match(r"sKVP:\d", ex_data[i]):
                self.kvp.append(ex_data[i]) #KVP
            elif re.match(r"Current:\d", ex_data[i]):
                self.xrtc.append(ex_data[i]) #X-Ray Tube Current (mA)
            elif ex_data[i-1] == "mAExposure" and re.match(r"Time:\d", ex_data[i]):
                self.ext.append(ex_data[i]) #Exposure Time (ms)
            elif re.match(r"msExposure:\d", ex_data[i]):
                self.exp.append(ex_data[i]) #Exposure (uA*s)
            elif ex_data[i-1] == "msPulse" and re.match(r"Width:\d", ex_data[i]):
                self.pulw.append(ex_data[i]) #Pulse Width
            elif re.match(r"\d+Irradiation",ex_data[i-1]) and re.match(r"Duration:\d",ex_data[i]):
                self.ird.append(ex_data[i]) #Irradiation Duration
            elif re.match(r"Detector:\d", ex_data[i]):
                self.dstd.append(ex_data[i]) #Distance Source to Detector (mm)
            elif re.match(r"Height:\d", ex_data[i]):
                self.cfh.append(ex_data[i]) #Collimated Field Height (mm)
            elif re.match(r"Width:\d", ex_data[i]):
                self.cfw.append(ex_data[i]) #Collimated Field Width (mm)
                n += 1
            elif ex_data[i-1] == "Thickness" and re.match(r"Maximum:\d",ex_data[i]):
                self.xrftm.append(ex_data[i]) #X-Ray Filter Thickness Maximum (mmCu)

        return [self.name, self.manuf, self.cont, self.obs, self.totaldata,self.distance, self.time, self.rpd, self.dap, self.fm, self.pet, self.drp, self.ppa, self.psa, self.xfm, self.cfa, self.xrftm, self.pr, self.kvp, self.xrtc, self.ext, self.exp, self.pulw, self.ird, self.dstd, self.cfh, self.cfw, n]



    def startpro(self,word_file):

        word_file_path =word_file
        # Extract data from the Word file
        ex_data = self.extract_data_from_pdf(word_file_path)

        self.ndata = self.newdata(ex_data)

        self.name, self.manuf, self.cont, self.obs, self.totaldata, self.distance, self.time, self.rpd, self.dap, self.pet, self.fm, self.drp, self.ppa, self.psa, self.xfm, self.cfa, self.xrftm, self.pr, self.kvp, self.xrtc, self.ext, self.exp, self.pulw, self.ird, self.dstd, self.cfh, self.cfw, self.events = self.ndata

     ###### make all list without words######
        self.name_id = [s.replace('Patient:', '').replace(')Study:Ortho/TraumaSeries:Radiation', '').replace('#', '').replace(',', '')
                        .replace(')Study:GeneralSeries:Radiation', '').replace(')Study:InterventionSeries:Radiation', '') for s in self.name]

        self.manufacturer = self.manuf[0].replace('InformationManufacturer:', '').replace('"', '')
        print(self.manufacturer)
        self.content = self.cont[0].replace('*', '').replace(',', '')

        self.obs[0] = self.obs[0].replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.observer = self.obs[0].split()
        self.total = [s.replace('Total:', '').replace('empty', '').replace('Reference', '').replace('Distance', '').replace('Total', '') for s in self.totaldata]

        self.total = [float(value) if (value is not None and value != '') else None for value in self.total]
        self.distance = [s.replace('Point:', '') for s in self.distance]
        self.time = [s.replace('Time:', '') for s in self.time]

        self.rpd = [s.replace('Deﬁnition:', '').replace("Definition:","").replace("cm","") for s in self.rpd]
        self.xrftm1 = [s.replace('Maximum:', '') for s in self.xrftm]
        self.xfm = [s.replace('Material:', '') for s in self.xfm]

        self.moddap = [item.split(' ')[0] for item in self.dap]
        self.dap1 = []
        self.pet1 = []
        self.drp1 = []
        self.ppa1 =[]
        self.psa1 =[]
        self.cfa1= []
        self.pr1= []
        self.xrtc1= []
        self.kvp1= []
        self.ext1= []
        self.pulw1= []
        self.exp1= []
        self.ird1= []
        self.dstd1= []
        self.cfh1= []
        self.cfw1= []
        convert_and_append(self.moddap, self.dap1)
        convert_and_append(self.pet, self.pet1)
        convert_and_append(self.drp, self.drp1)
        convert_and_append(self.ppa, self.ppa1)
        convert_and_append(self.psa, self.psa1)
        convert_and_append(self.cfa, self.cfa1)
        convert_and_append(self.pr, self.pr1)
        convert_and_append(self.xrtc, self.xrtc1)
        convert_and_append(self.kvp, self.kvp1)
        convert_and_append(self.ext, self.ext1)
        convert_and_append(self.pulw, self.pulw1)
        convert_and_append(self.exp, self.exp1)
        convert_and_append(self.ird, self.ird1)
        convert_and_append(self.dstd, self.dstd1)
        convert_and_append(self.cfh, self.cfh1)
        convert_and_append(self.cfw, self.cfw1)

        # Create a pandas DataFrame from the extracted data
        self.all_data = { "Dose Area Product (Gym2)":self.dap1,"Dose (RP) (Gy)":self.drp1, "X-Ray Filter Material":self.xfm,
                           "Positioner Primary Angle (deg)":self.ppa1,"Positioner Secondary Angle (deg)":self.psa1,'Collimated Field Area (m2)':self.cfa1,
                           'X-Ray Filter Thickness Maximum (mmCu)':self.xrftm1,"Pulse Rate (pulse/s)":self.pr1,'KVP':self.kvp1,'X-Ray Tube Current (mA)':self.xrtc1,
                           'Exposure Time (ms)':self.ext1,'Pulse Width (ms)':self.pulw1,'Exposure (uA.s)':self.exp1,'Irradiation Duration (s)':self.ird1,'Distance Source to Detector (mm)':self.dstd1,
                           'Collimated Field Height (mm)':self.cfh1,'Collimated Field Width (mm)':self.cfw1}


        # Use regular expressions to extract numbers from mixed elements
        max_length = max(len(self.all_data[col]) for col in self.all_data)
        for col in self.all_data:
            self.all_data[col] += [np.nan] * (max_length - len(self.all_data[col]))

        self.df = pd.DataFrame(self.all_data)
        for i in range(0,self.events):
            self.df = self.df.rename(index={i: "Event {0}".format(i+1)})


        self.df = self.df.rename_axis('Irradiation Event X-Ray Data of '+ self.name_id[0]+" "+self.name_id[1])
        print(self.rpd)
        self.data_total = {"ID": self.name_id[2],"Manufacturer":self.manufacturer,"Content Date":self.content ,"Observer":self.observer[0], "Dose Area Product Total (Gym2)": self.total[0],
                           "Dose (RP) Total (Gy)":self.total[1],"Distance Source to Reference Poit (mm)":self.distance[0],"Fluoro Dose Area Product Total (Gym2)":self.total[3],
                                        "Fluoro Dose (RP) Total (Gy)":self.total[4],	"Total Fluoro Time (s)":self.time[0],"Acquisition Dose Area Product Total (Gym2)":self.total[4],
                                        "Acquisition Dose (RP) Total (Gy)":self.total[0],"Reference Point Definition (cm)":self.rpd[0],	"Total Acquisition Time (s)":self.total[2]}
        self.dft = pd.DataFrame(self.data_total, index=[self.name_id[0]+" "+self.name_id[1]])
        self.dft = self.dft.rename_axis('Accumulated X-Ray Dose Data')

        self.person_data = [self.name_id[0]+" "+self.name_id[1],self.name_id[2],self.manufacturer,self.content[0],self.observer[0],self.events]

        self.dfper = pd.DataFrame(self.person_data, index=['Patient name', 'ID', 'Manufacturer', 'Content Date', 'Person Observer Name', 'Number of irradiation events'], columns=[""])

        '''print(self.df)
        print(self.dft)
        print(self.dfper)'''


        return self.df, self.dft, self.dfper, self.name_id[1], self.name_id[2]
        # Output directory where Excel files will be saved


    def get_dataframes(self):
        return self.df, self.dft, self.dfper, self.name_id[1], self.name_id[2]

def convert_and_append(source_list, target_list):
    result_list = []
    for item in source_list:
        numeric_sequences = re.findall(r'\b\d+(?:\.\d+)?\b', item)

        # Add each numeric sequence to the result
        for seq in numeric_sequences:
            result_list.append(seq)

    result_list = [float(num) if '.' in num else int(num) for num in result_list]
    target_list.extend(result_list)