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
        self.patthick = []
        self.time =[]
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
        n = -1
        self.ext = []
        self.exp= []
        self.pulw = []
        self.ird = []
        self.dstd = []
        self.cfh = []
        self.cfw = []
        print(ex_data)
        for i in range(0, len(ex_data)):
            if ex_data [i] == "Event" and ex_data[i+1] == "X-Ray" :
                n += 1
                self.patthick.append(None)
                self.dap.append(None)
                self.fm.append(None)
                self.pet.append(None)
                self.drp.append(None)
                self.ppa.append(None)
                self.psa.append(None)
                self.xfm.append(None)
                self.cfa.append(None)
                self.xrftm.append(None)
                self.pr.append(None)
                self.kvp.append(None)
                self.xrtc.append(None)
                self.ext.append(None)
                self.exp.append(None)
                self.pulw.append(None)
                self.ird.append(None)
                self.dstd.append(None)
                self.cfh.append(None)
                self.cfw.append(None)
            if ex_data [i] == "Person":
                if ex_data[i+1] == "Observer" and ex_data[i+2] == "Name":
                    self.obs.append(ex_data[i+4])
                    if ex_data[i+4] == "DR.":
                        self.obs[0] = ex_data[i+4] +" "+ ex_data[i+5]
            if re.match(r'Manufacturer', ex_data[i]):
                self.manuf.append(ex_data[i+2])
            elif re.match(r'Total:+\d|Total:\w',ex_data[i]):
                self.totaldata.append(ex_data[i]) #Dose Area Product Total, Dose (RP) Total,
                # Fluoro Dose Area Product Total, Fluoro Dose (RP) Total, Acquisition Dose Area Product Total
            elif re.match(r'Point:+\d',ex_data[i]):
                self.distance.append(ex_data[i]) #Distance Source to Reference Point
            elif re.match(r'Time:+\d',ex_data[i]):
                if ex_data[i - 1] == "Flouro" or ex_data[i - 1] == "Acquisition":
                    self.time.append(ex_data[i]) # Total Fluoro Time (s), Total Acquisition Time,
                # Exposure time
                if ex_data[i - 1] == "mAExposure":
                    self.ext [n] = ex_data [i]  # Exposure Time (ms)
            elif re.match(r'Definition:+\d|Deﬁnition:+\d',ex_data[i]):
                self.rpd.append(ex_data[i]) #Reference Point Definition (cm)
            elif re.match(r"Product:+\d",ex_data[i]):
                self.dap [n] = ex_data [i] #Dose area product
            elif re.match(r"Thickness:+\d",ex_data[i]) and re.match(r"Equivalent",ex_data[i-1]):
                self.patthick [n] = ex_data [i] #Patient Equivalent Thickness (mm)
            elif re.match(r"\(RP\):\d",ex_data[i]):
                self.drp [n] = ex_data [i] # Dose (RP) (Gy)
            elif re.match(r"Angle:-?\d+(\.\d+)?|Angle:?\d+(\.\d+)?", ex_data[i]):
                if "Primary" == ex_data[i-1]:
                    self.ppa [n] = ex_data [i]  # Positioner Primary Angle (deg)
                elif "Secondary" == ex_data[i-1]:
                    self.psa [n] = ex_data [i]  # Positioner Secondary Angle (deg)
            elif re.match(r"Material:\w",ex_data[i]):
                self.xfm [n] = ex_data [i] #X-Ray Filter Material
            elif re.match(r"Mode:\w",ex_data[i]):
                self.fm [n] = ex_data [i] # Fluoro Mode
            elif re.match(r"Area:\d",ex_data[i]) and "Field" == ex_data[i-1]:
                self.cfa [n] = ex_data [i] # Collimated Field Area (m2)
            elif re.match(r"Rate:\d",ex_data[i]):
                self.pr [n] = ex_data [i] #Pulse Rate (pulse/s)
            elif re.match(r"sKVP:\d|\dKVP:\d", ex_data[i]):
                self.kvp [n] = ex_data [i] #KVP
            elif re.match(r"Current:\d", ex_data[i]):
                self.xrtc [n] = ex_data [i] #X-Ray Tube Current (mA)
            elif re.match(r"msExposure:\d|sExposure:\d|Exposure:\d", ex_data[i]):
                self.exp [n] = ex_data [i] #Exposure (uA*s)
            elif re.match(r"Width:\d", ex_data[i]) and ex_data[i-1] == "msPulse" :
                self.pulw [n] = ex_data [i] #Pulse Width
            elif re.match(r"\d+Irradiation|msIrradiation",ex_data[i-1]) and re.match(r"Duration:\d",ex_data[i]):
                self.ird [n] = ex_data [i] #Irradiation Duration
            elif re.match(r"Detector:\d", ex_data[i]):
                self.dstd [n] = ex_data [i] #Distance Source to Detector (mm)
            elif re.match(r"Height:\d", ex_data[i])and re.match(r"Field",ex_data[i-1]):
                self.cfh [n] = ex_data [i] #Collimated Field Height (mm)
            elif re.match(r"Width:\d", ex_data[i]) and re.match(r"Field",ex_data[i-1]):
                self.cfw [n] = ex_data [i] #Collimated Field Width (mm)
            elif re.match(r"Maximum:\d",ex_data[i]) and ex_data[i-1] == "Thickness":
                self.xrftm [n] = ex_data [i] #X-Ray Filter Thickness Maximum (mmCu)

        return [self.name, self.manuf, self.cont, self.obs, self.totaldata,self.distance, self.time, self.rpd, self.dap,self.patthick, self.fm, self.pet, self.drp, self.ppa, self.psa, self.xfm, self.cfa, self.xrftm, self.pr, self.kvp, self.xrtc, self.ext, self.exp, self.pulw, self.ird, self.dstd, self.cfh, self.cfw, n]



    def startpro(self,word_file):

        word_file_path =word_file
        # Extract data from the Word file
        ex_data = self.extract_data_from_pdf(word_file_path)

        self.ndata = self.newdata(ex_data)

        self.name, self.manuf, self.cont, self.obs, self.totaldata, self.distance, self.time, self.rpd, self.dap,self.patthick, self.pet, self.fm, self.drp, self.ppa, self.psa, self.xfm, self.cfa, self.xrftm, self.pr, self.kvp, self.xrtc, self.ext, self.exp, self.pulw, self.ird, self.dstd, self.cfh, self.cfw, self.events = self.ndata

     ###### make all list without words######
        self.name_id = [s.replace('Patient:', '').replace(')Study:Ortho/TraumaSeries:Radiation', '').replace('#', '').replace(',', '')
                        .replace(')Study:GeneralSeries:Radiation', '').replace(')Study:InterventionSeries:Radiation', '').replace(')Study:Coronary^Diagnostic', '')
                        for s in self.name]
        self.obs = [s.replace("[","").replace("]","").replace("'","") for s in self.obs]
        self.manufacturer = self.manuf[0].replace('InformationManufacturer:', '').replace('"', '')
        self.content = self.cont[0].replace('*', '').replace(',', '')
        self.total = [s.replace('Total:', '').replace('empty', '').replace('Reference', '').replace('Distance', '').replace('Total', '') for s in self.totaldata]
        self.total = [float(value) if (value is not None and value != '') else None for value in self.total]
        self.distance = [s.replace('Point:', '') for s in self.distance]
        self.time = [s.replace('Time:', '') for s in self.time]
        self.rpd = [s.replace('Deﬁnition:', '').replace("Definition:","").replace("cm","") for s in self.rpd]

        self.xrftm1 = [s.replace('Maximum:', '') if s is not None else None for s in self.xrftm]

        self.xfm = [s.replace('Material:', '') if s is not None else None for s in self.xfm]

        self.moddap = [item.split(' ')[0] if item is not None else None for item in self.dap]

        self.dap1 = []
        self.patthick1 =[]
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
        convert_and_append(self.patthick, self.patthick1)
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
        self.all_data = { "Dose Area Product (Gym2)":self.dap1,"Patient Equivalent Thickness (mm)":self.patthick1,"Dose (RP) (Gy)":self.drp1, "X-Ray Filter Material":self.xfm,
                           "Positioner Primary Angle (deg)":self.ppa1,"Positioner Secondary Angle (deg)":self.psa1,'Collimated Field Area (m2)':self.cfa1,
                           'X-Ray Filter Thickness Maximum (mmCu)':self.xrftm1,"Pulse Rate (pulse/s)":self.pr1,'KVP':self.kvp1,'X-Ray Tube Current (mA)':self.xrtc1,
                           'Exposure Time (ms)':self.ext1,'Pulse Width (ms)':self.pulw1,'Exposure (uA.s)':self.exp1,'Irradiation Duration (s)':self.ird1,'Distance Source to Detector (mm)':self.dstd1,
                           'Collimated Field Height (mm)':self.cfh1,'Collimated Field Width (mm)':self.cfw1}


        # Use regular expressions to extract numbers from mixed elements
        max_length = max(len(self.all_data[col]) for col in self.all_data)
        for col in self.all_data:
            self.all_data[col] += [np.nan] * (max_length - len(self.all_data[col]))

        self.df = pd.DataFrame(self.all_data)
        for i in range(0,self.events+1):
            self.df = self.df.rename(index={i: "Event {0}".format(i+1)})


        self.df = self.df.rename_axis('Irradiation Event X-Ray Data of '+ self.name_id[0]+" "+self.name_id[1])
        self.data_total = {"ID": self.name_id[2],"Manufacturer":self.manufacturer,"Content Date":self.content ,"Observer":self.obs, "Dose Area Product Total (Gym2)": self.total[0],
                           "Dose (RP) Total (Gy)":self.total[1],"Distance Source to Reference Poit (mm)":self.distance[0],"Fluoro Dose Area Product Total (Gym2)":self.total[3],
                                        "Fluoro Dose (RP) Total (Gy)":self.total[4],	"Total Fluoro Time (s)":self.time[0],"Acquisition Dose Area Product Total (Gym2)":self.total[4],
                                        "Acquisition Dose (RP) Total (Gy)":self.total[0],"Reference Point Definition (cm)":self.rpd[0],	"Total Acquisition Time (s)":self.total[2]}
        self.dft = pd.DataFrame(self.data_total, index=[self.name_id[0]+" "+self.name_id[1]])
        self.dft = self.dft.rename_axis('Accumulated X-Ray Dose Data')

        self.person_data = [self.name_id[0]+" "+self.name_id[1],self.name_id[2],self.manufacturer,self.content[0],self.obs ,self.events+1]

        self.dfper = pd.DataFrame(self.person_data, index=['Patient name', 'ID', 'Manufacturer', 'Content Date', 'Person Observer Name', 'Number of irradiation events'], columns=[""])

        '''print(self.df)
        print(self.dft)
        print(self.dfper)'''


        return self.df, self.dft, self.dfper, self.name_id[1], self.name_id[2]
        # Output directory where Excel files will be saved


    def get_dataframes(self):
        return self.df, self.dft, self.dfper, self.name_id[1], self.name_id[2]

import re

def convert_and_append(source_list, target_list):
    result_list = []
    for item in source_list:
        if item is None:
            result_list.append(None)
        elif isinstance(item, str):
            numeric_sequences = re.findall(r'[-+]?\b\d+(?:\.\d+)?(?:[eE][-+]?\d+)?\b', item)

            # Add each numeric sequence to the result
            for seq in numeric_sequences:
                result_list.append(float(seq))
        else:
            result_list.append(None)  # Handle non-string, non-None values

    target_list.extend(result_list)
