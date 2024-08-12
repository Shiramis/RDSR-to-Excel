import pandas as pd
import numpy as np
import PyPDF2
import datetime

class make_excel ():

    def extract_data_from_pdf(self, file_path):
        data = []
        ο = 0
        with open(file_path, 'rb') as file:
            try:
                pdf_reader = PyPDF2.PdfReader(file)  # Continue with your processing code here
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    data.extend(page.extract_text().split())
                    ο = 1
            except PyPDF2.errors.PdfReadError as e:
                print(f"Error reading PDF file: {e}")

        return data, ο

    def newdata(self,ex_data):
        print(ex_data)
        self.name = []
        i = 0
        while not re.search(r'Series', str(ex_data[i])):
            self.name.append(ex_data[i])
            i += 1
        else:
            self.name.append(ex_data[i])
        print(self.name)
        self.birth =[]
        self.manuf = []
        self.cont = []
        self.cont.extend(["N/A","N/A"])
        self.obs = []
        self.patthick = []
        self.time =[]
        self.time = ["N/A"]* 5
        self.totaldata = []
        self.totaldata = ["N/A"] * 6
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
        self.iso = []
        self.defin = []

        for i in range(len(ex_data)):

            if ex_data [i] == "Event" and ex_data[i+1] == "X-Ray" :
                n += 1

                self.patthick.append("N/A")
                self.dap.append("N/A")
                self.fm.append("N/A")
                self.pet.append("N/A")
                self.drp.append("N/A")
                self.ppa.append("N/A")
                self.psa.append("N/A")
                self.xfm.append("N/A")
                self.cfa.append("N/A")
                self.xrftm.append("N/A")
                self.pr.append("N/A")
                self.kvp.append("N/A")
                self.xrtc.append("N/A")
                self.ext.append("N/A")
                self.exp.append("N/A")
                self.pulw.append("N/A")
                self.ird.append("N/A")
                self.dstd.append("N/A")
                self.cfh.append("N/A")
                self.cfw.append("N/A")
                self.iso.append("N/A")
                self.distance.append("N/A")
            if ex_data [i] == "Person":
                if ex_data[i+1] == "Observer" and ex_data[i+2] == "Name":
                    self.obs.append(ex_data[i+4])
                    if ex_data[i+4] == "DR.":
                        self.obs[0] = ex_data[i+4] +" "+ ex_data[i+5]
            if re.match(r'\*\d|\*+\d', ex_data[i]):
                self.birth.append(ex_data[i])
            elif re.match(r'Date/Time:\d|Started:\d', ex_data[i]):
                self.cont[0] = ex_data[i]
                self.cont[1] = ex_data[i+1]
            elif re.match(r'Manufacturer', ex_data[i]):
                self.manuf.append(ex_data[i+2])
            elif re.match(r'Model', ex_data[i]) and re.match(r'Observer', ex_data[i-1]):
                self.manuf[0] += " " + ex_data[i+3]
            elif re.match(r'Total:+\d|Total:\w|Total:emptyReference|Total:emptyTotal|Total:emptyDistance',ex_data[i]):
                index = self.totaldata.index("N/A")
                if ex_data[i] == "Total:emptyReference" or ex_data [i] == "Total:emptyDistance"\
                        or ex_data [i] == 'Total:emptyTotal':
                    self.totaldata [index] = 0
                else:
                    self.totaldata[index] = ex_data[i] #Dose Area Product Total, Dose (RP) Total,
                # Fluoro Dose Area Product Total, Fluoro Dose (RP) Total, Acquisition Dose Area Product Total
                #Acquisition Dose (RP) Total
            elif re.match(r'Point:+\d|Point:\d|Point:',ex_data[i]):
                if n>=0:
                    if ex_data [i] == "Point:":
                        self.distance[n] = ex_data[i+1]
                    else:
                        self.distance[n] = ex_data[i]#Distance Source to Reference Point

            elif re.match(r'Time:+\d|Time:\d|Time:',ex_data[i]):
                if ex_data[i - 1] == "Fluoro" or ex_data[i - 1] == "Acquisition":
                    index1 = self.time.index("N/A")
                    self.time[index1] = ex_data[i] # Total Fluoro Time (s), Total Acquisition Time
                if ex_data[i - 1] == "mAExposure":
                    if ex_data[i] == "Time:":
                        self.ext[n] = ex_data[i+1]
                    else:
                        self.ext [n] = ex_data [i]  # Exposure Time (ms)
            elif re.match(r'Definition:+\d|Deﬁnition:+\d|Definition:|Deﬁnition:',ex_data[i]):
                if ex_data[i] == "Definition:" or ex_data[i] == 'Deﬁnition':
                    self.rpd.append(ex_data[i+1])
                else:
                    self.rpd.append(ex_data[i]) #Reference Point Definition (cm)
            elif re.match(r"Product:+\d|\wProduct:\d|Product:",ex_data[i]):
                if ex_data [i] == "Product:":
                    self.dap [n] = ex_data [i+1]
                elif ex_data [i] =='Product:emptyDose':
                    self.dap[n] = "N/A"
                else:
                    self.dap [n] = ex_data [i] #Dose area product DAP
            elif re.match(r"Thickness:+\d|Thickness:\d|Thickness:",ex_data[i]) and re.match(r"Equivalent",ex_data[i-1]):
                if ex_data [i] =="Thickness:":
                    self.patthick [n] = ex_data [i+1]
                else:
                    self.patthick [n] = ex_data [i] #Patient Equivalent Thickness (mm)
            elif re.match(r"\(RP\):\d|\(RP\) :\d|\(RP\):|\(RP\):emptyPositioner",ex_data[i]):
                if re.match(r"\d+|\d", ex_data[i + 1]):
                    self.drp[n] = ex_data [i+1]
                elif ex_data [i] == r"(RP):emptyPositioner":
                    self.drp[n] = "N/A"
                else:
                    self.drp [n] = ex_data [i] # Dose (RP) (Gy)
            elif re.match(r"Angle:-?\d+(\.\d+)?|Angle:?\d+(\.\d+)?|Angle:", ex_data[i]):
                if "Primary" == ex_data[i-1]:
                    if ex_data[i]=="Angle:":
                        self.ppa[n] = ex_data[i+1]
                    else:
                        self.ppa [n] = ex_data [i]  # Positioner Primary Angle (deg)
                elif "Secondary" == ex_data[i-1]:
                    if ex_data[i]=="Angle:":
                        self.psa[n] = ex_data[i+1]
                    else:
                        self.psa [n] = ex_data [i]  # Positioner Secondary Angle (deg)
            elif re.match(r"Material:\w|Material:|Material:\w+",ex_data[i]):
                if ex_data[i] == "Material:":
                    self.xfm [n] = ex_data [i+1]
                else:
                    self.xfm [n] = ex_data [i] #X-Ray Filter Material
            elif re.match(r"Mode:\w|Mode:",ex_data[i]):
                if ex_data[i] == "Mode:":
                    self.fm[n] = ex_data[i+1]
                else:
                    self.fm [n] = ex_data [i] # Fluoro Mode
            elif re.match(r"Area:\d|Area:",ex_data[i]) and "Field" == ex_data[i-1]:
                if ex_data[i] == "Area:":
                    self.cfa [n] = ex_data [i+1]
                else:
                    self.cfa [n] = ex_data [i] # Collimated Field Area (m2)
            elif re.match(r"Rate:\d|Rate:",ex_data[i]):
                if ex_data[i] == "Rate:":
                    self.pr [n] = ex_data[i+1]
                else:
                    self.pr [n] = ex_data [i] #Pulse Rate (pulse/s)
            elif re.match(r"sKVP:\d|\dKVP:\d|KVP:|sKVP:|\wKVP:\d", ex_data[i]):
                if ex_data [i] == "sKVP:" or ex_data == "KVP:":
                    self.kvp [n] =ex_data [i+1]
                else:
                    self.kvp [n] = ex_data [i] #KVP
            elif re.match(r"Current:\d|Current:", ex_data[i]):
                if ex_data [i] == "Current:":
                    self.xrtc [n] = ex_data [i+1]
                else:
                    self.xrtc [n] = ex_data [i] #X-Ray Tube Current (mA)
            elif re.match(r"msExposure:\d|sExposure:\d|Exposure:\d|Exposure:|\wExposure:\d", ex_data[i]):
                if ex_data [i] == "Exposure:":
                    self.exp [n] = ex_data [i+1]
                else:
                    self.exp [n] = ex_data [i] #Exposure (uA*s)
            elif re.match(r"Width:\d|Width:", ex_data[i]) and ex_data[i-1] == "msPulse" :
                if ex_data [i] == "Width:":
                    self.pulw [n] = ex_data[i+1]
                else:
                    self.pulw [n] = ex_data [i] #Pulse Width
            elif re.match(r"Duration:\d|Duration:",ex_data[i]) :
                if self.ird[n].strip() == "N/A":
                    if ex_data[i] == "Duration:":
                        self.ird [n] = ex_data[i+1]
                    else:
                        self.ird [n] = ex_data [i] #Irradiation Duration
            elif re.match(r"Detector:\d|Detector:", ex_data[i]):
                if ex_data [i] == "Detector:":
                    self.dstd [n] = ex_data[i+1]
                else:
                    self.dstd [n] = ex_data [i] #Distance Source to Detector (mm)
            elif re.match(r"Height:\d|Height:", ex_data[i]) and re.match(r"Field",ex_data[i-1]):
                if ex_data[i] == "Height:":
                    self.cfh [n] =ex_data [i+1]
                else:
                    self.cfh [n] = ex_data [i] #Collimated Field Height (mm)
            elif re.match(r"Width:\d|Width:", ex_data[i]) and re.match(r"Field",ex_data[i-1]):
                if ex_data [i] == "Width:":
                    self.cfw [n] = ex_data[i+1]
                else:
                    self.cfw [n] = ex_data [i] #Collimated Field Width (mm)
            elif re.match(r"Maximum:\d|Maximum:",ex_data[i]) and ex_data[i-1] == "Thickness":
                if ex_data [i] == "Maximum:":
                    self.xrftm[n] = ex_data [i+1]
                else:
                    self.xrftm [n] = ex_data [i] #X-Ray Filter Thickness Maximum (mmCu)
            elif re.match(r"Isocenter:\d|Isocenter:",ex_data[i]):
                if ex_data[i] == "Isocenter:":
                    self.iso [n] = ex_data [i+1]
                else:
                    self.iso [n] = ex_data [i] # Distance Source to Isocenter (mm)
        return n

    def startpro(self,word_file, index):

        word_file_path =word_file
        all = self.extract_data_from_pdf(word_file_path)
        ex_data, o = all
        if o == 1:
            self.ndata = self.newdata(ex_data)
            self.events = self.ndata
         ###### make all list without words######
            i = 0
            while not re.search(r'Study', str(self.name[i])):
                if re.match(r"\(+\w",self.name[i]):
                    self.sex = self.name[i]
                self.start_index = self.name[i]
                indexs = i
                i += 1
            self.study = ""
            for i in range(indexs+1,len(self.name)):
                self.study = self.study+" "+self.name[i]
            self.study = re.sub(r'^.*Study:', '', self.study)
            self.study = re.sub(r'Series:+\w+', '', self.study)
            self.study =self.study.replace("^", " ").replace("/"," ")
            self.sex = self.sex.replace('(','').replace(',','')

            self.name_id = [s.replace('Patient:', '').replace(')Study:Ortho/TraumaSeries:Radiation', '').replace('#', '').replace(',', '')
                            .replace(')Study:GeneralSeries:Radiation', '').replace(')Study:InterventionSeries:Radiation', '').replace(')Study:Coronary^Diagnostic', '')
                            for s in self.name]
            self.obs = [s.replace("[","").replace("]","").replace("'","") for s in self.obs]
            self.manufacturer = self.manuf[0].replace('InformationManufacturer:', '').replace('"', '')
            if self.cont[0] != "N/A":
                self.content = self.cont[0].replace('*', '').replace(',', '').replace('Date/Time:', '').replace('Started:','')
                self.contime = self.cont[1].replace('X-Ray','').replace('Irradiation','')

            if self.totaldata != "N/A":
                self.total = [s.replace('Total:', '').replace('empty', '').replace('Reference', '').replace('Distance', '').replace('Total', '')
                          .replace('Fluoro', '') if isinstance(s, str) else s for s in self.totaldata]
                self.total = [float(value) if (value != "N/A" and value != '') else "N/A" for value in self.total]

            self.distance = [s.replace('Point:', '') if s != "N/A" else "N/A" for s in self.distance]
            if self.time != "N/A":
                self.time = [s.replace('Time:', '') if isinstance(s, str) else s for s in self.time]

            self.rpd = [s.replace('Deﬁnition:', '').replace("Definition:","").replace("cm","") for s in self.rpd]

            self.xrftm1 = [s.replace('Maximum:', '') if s != "N/A" else "N/A" for s in self.xrftm]

            self.xfm = [s.replace('Material:', '') if s != "N/A" else "N/A" for s in self.xfm]

            self.iso = [s.replace('Isocenter:', '') if s != "N/A" else "N/A" for s in self.iso]
            self.defin = [s.replace('Point:', '') if s != "N/A" else "N/A" for s in self.defin]


            self.birth = self.birth[0].split(',')[0][1:]
            year, month, day = map(int, self.birth.split("-"))
            today = datetime.date.today()
            self.age = today.year - year - ((today.month, today.day) < (month, day))

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
            self.iso1 =[]
            self.defin1 = []

            convert_and_append(self.dap, self.dap1)
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
            convert_and_append(self.iso, self.iso1)
            convert_and_append(self.defin, self.defin1)
            self.all_data = { "Dose Area Product (Gym\u00b2)":self.dap1,"Patient Equivalent Thickness (mm)":self.patthick1,"Dose (RP) (Gy)":self.drp1,
                                   "Positioner Primary Angle (deg)":self.ppa1,"Positioner Secondary Angle (deg)":self.psa1,'Collimated Field Area (m\u00b2)':self.cfa1,
                                  'Collimated Field Height (mm)':self.cfh1,'Collimated Field Width (mm)':self.cfw1,
                                  "X-Ray Filter Material":self.xfm,
                                   'X-Ray Filter Thickness Maximum (mmCu)':self.xrftm1,"Pulse Rate (pulse/s)":self.pr1,
                                  'KVP':self.kvp1,'X-Ray Tube Current (mA)':self.xrtc1,
                                   'Exposure Time (ms)':self.ext1,'Pulse Width (ms)':self.pulw1,'Irradiation Duration (s)':self.ird1,
                                  'Exposure (uA.s)':self.exp1,
                                   'Distance Source to Detector (mm)':self.dstd1,"Distance Source to Isocenter (mm)": self.iso1,
                                  "Distance Source to Reference Point (mm)":self.distance}

            max_length = max(len(self.all_data[col]) for col in self.all_data)
            for col in self.all_data:
                self.all_data[col] += [np.nan] * (max_length - len(self.all_data[col]))

            self.df = pd.DataFrame(self.all_data)
            for i in range(0,self.events+1):
                self.df = self.df.rename(index={i: "Event {0}".format(i+1)})
            #name = self.name_id[0]+" "+self.name_id[1]
            name = f"Patient {index}"
            #ID = self.name_id[2]
            ID = f"ID {index}"
            self.obs = f"Observer {index}"

            self.df = self.df.rename_axis('Irradiation Event X-Ray Data of '+ name )
            self.data_total = {"Patient ID": ID, "Dose Area Product Total (Gym\u00b2)": self.total[0],
                               "Dose (RP) Total (Gy)":self.total[1],"Fluoro Dose Area Product Total (Gym\u00b2)":self.total[2],
                                            "Fluoro Dose (RP) Total (Gy)":self.total[3],	"Total Fluoro Time (s)":self.time[0],"Acquisition Dose Area Product Total (Gym\u00b2)":self.total[4],
                                            "Acquisition Dose (RP) Total (Gy)":self.total[5],	"Total Acquisition Time (s)":self.time[1],"Reference Point Definition (cm)": self.rpd[0]}
            self.dft = pd.DataFrame(self.data_total, index=[name])
            self.dft = self.dft.rename_axis('Patient Name')

            self.person_data = [self.name_id[0]+" "+self.name_id[1],self.name_id[2],self.sex, self.age, self.study, self.manufacturer,
                                self.content,self.contime, self.obs ,self.events+1]

            self.dfper = pd.DataFrame(self.person_data, index=['Patient Name', 'Patient ID','Gender','Age (years)', 'Study Type', 'Manufacturer', 'Content Time', 'Content Date', 'Person Observer Name',
                                         'Number of irradiation events'], columns=[""])
            self.individual = {" Patient ID": ID,"Gender": self.sex,"Age (years)":self.age,"Study Type":self.study,
                               "Manufacturer": self.manufacturer, "Content Date": self.content, "Content Time": self.contime,
                                   "Person Observer Name": self.obs}
            self.dfin = pd.DataFrame(self.individual, index=[name])
            self.dfin = self.dfin.rename_axis('Patient Name')
        else:
            self.name_id = []
            self.name_id = ["N/A"]* 5
            self.person_data =[]
            self.person_data = ["N/A"]*10
            self.all_data = {"Dose Area Product (Gym\u00b2)": ["N/A"], "Dose (RP) (Gy)": ["N/A"],
                             "Positioner Primary Angle (deg)": ["N/A"], "Positioner Secondary Angle (deg)": ["N/A"],
                             "X-Ray Filter Material": ["N/A"], 'X-Ray Filter Thickness Maximum (mmCu)': ["N/A"],
                             "Pulse Rate (pulse/s)": ["N/A"], 'Irradiation Duration (s)': ["N/A"], 'KVP': ["N/A"],
                             'X-Ray Tube Current (mA)': ["N/A"], 'Exposure Time (ms)': ["N/A"],
                             'Pulse Width (ms)': ["N/A"], 'Exposure (uA.s)': ["N/A"],
                             'Collimated Field Area (m\u00b2)': ["N/A"], 'Collimated Field Height (mm)': ["N/A"],
                             'Collimated Field Width (mm)': ["N/A"], 'Distance Source to Detector (mm)': ["N/A"]}
            self.data_total = {"Patient ID": ["N/A"], "Manufacturer": ["N/A"], "Content Date": ["N/A"],
                               "Person Observer Name": ["N/A"], "Dose Area Product Total (Gym\u00b2)": ["N/A"],
                               "Dose (RP) Total (Gy)": ["N/A"],
                               "Fluoro Dose Area Product Total (Gym\u00b2)": ["N/A"],
                               "Fluoro Dose (RP) Total (Gy)": ["N/A"], "Total Fluoro Time (s)": ["N/A"],
                               "Acquisition Dose Area Product Total (Gym\u00b2)": ["N/A"],
                               "Acquisition Dose (RP) Total (Gy)": ["N/A"],
                               "Reference Point Definition (cm)": ["N/A"],
                               "Total Acquisition Time (s)": ["N/A"]}
            self.individual = {" Patient ID": ["N/A"], "Gender": ["N/A"], "Age (years)": ["N/A"], "Study Type": ["N/A"],
                               "Manufacturer": ["N/A"], "Content Date": ["N/A"],
                               "Content Time": ["N/A"], "Person Observer Name": ["N/A"]}
            name = f"Patient {index}"
            self.dft = pd.DataFrame(self.data_total, index=[name])
            self.df = pd.DataFrame(self.all_data)
            self.dfin = pd.DataFrame(self.individual, index=[name])
        return self.df, self.dft,self.individual,self.dfin, self.person_data, self.name_id[1], self.name_id[2]
    def get_dataframes(self):
        return self.df, self.dft, self.dfper, self.name_id[1], self.name_id[2]

import re

def convert_and_append(source_list, target_list):
    result_list = []
    for item in source_list:
        if item == "N/A":
            result_list.append("N/A")
        elif isinstance(item, str):
            numeric_sequences = re.findall(r'[-+]?\b\d+(?:\.\d+)?(?:[eE][-+]?\d+)?\b', item)
            for seq in numeric_sequences:
                result_list.append(float(seq))
        else:
            result_list.append("N/A")

    target_list.extend(result_list)
