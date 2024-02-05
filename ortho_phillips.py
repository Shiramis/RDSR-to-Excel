import pandas as pd
import re
import numpy as np
import PyPDF2

class make_excelp ():

    def extract_text_from_pdf(self,pdf_file_path):
        # Initialize an empty list to store the extracted data
        self.text = []
        with open(pdf_file_path, "rb") as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                self.text.extend(page.extract_text().split())
        return self.text

    def newdata(self,ex_data):
        self.name = []
        self.name.append(ex_data[0])
        self.name.append(ex_data[1])
        self.name.append(ex_data[4])
        self.id = []
        self.manuf = []
        self.cont = []
        self.obs = []
        self.obs.append(ex_data[10])

        self.totaldata = []
        self.rpd = []
        self.dap = []
        self.drp = []
        self.ppa = []
        self.psa = []
        self.xfm = []
        self.xft = []
        self.cfa = []
        self.xrftm = []
        n = 0
        self.ird = []
        self.dstd = []

        for i in range(0, len(ex_data)):
            if i<10:
                if re.findall(r'#',ex_data[i]) == ['#']:
                    self.id.append(ex_data[i])

            if ex_data[i] == "Information":
                self.manuf.append(ex_data[i+1])

            if ex_data[i] == 'Party' and ex_data[i+1] =='responsible'and ex_data[i+2] == 'for'and ex_data[i+3] == 'servicing'and ex_data[i+4] == 'the'and ex_data[i+5] == 'device':
                self.totaldata.append(ex_data[i + 6])# (1) Dose Area Product Total
                self.totaldata.append(ex_data[i + 7])# (2) Dose RP Total
                self.totaldata.append(ex_data[i + 9])# (3) Fluoro Dose Area Product Total
                self.totaldata.append(ex_data[i + 11])# (4) Fluoro Dose (RP) Total
                self.totaldata.append(ex_data[i + 13])# (5) Total Fluoro Time
                self.totaldata.append(ex_data[i + 15])# (6) Acquisition Dose Area Product Total
                self.totaldata.append(ex_data[i + 16])# (7) Acquisition Dose (RP) Total
                self.totaldata.append(ex_data[i + 18])# (8) Total Acuisition time
                self.totaldata.append(ex_data[i + 22])# Distance
                print(self.totaldata)
            if ex_data [i] == 'Flag:UNVERIFIEDContent':
                self.cont.append(ex_data[i+1])
            if 'Observer' == ex_data[i] and ex_data[i - 1] == 'Device' and ex_data[i + 1] == 'Name':
                self.obs.append(ex_data[i+3])

            if ex_data[i] == "Orthopaedics":
                n +=1
                self.xrftm.append(None)
                self.dap.append(ex_data[i + 8])# Dose Area Product
                self.drp.append(ex_data[i + 10])# Dose RP
                self.cfa.append(ex_data[i + 11])# Collimated Field Area

            if ex_data[i] == "ﬁlter" and ex_data[i+1]=="Copper":
                self.xfm.append(ex_data[i+3]+" "+ex_data[i+4])
                self.xft.append(ex_data[i + 5])
                self.ird.append(ex_data[i + 7])

        return [self.name,self.id, self.manuf, self.cont, self.obs, self.totaldata, self.rpd, self.dap, self.drp, self.ppa, self.psa, self.xfm, self.cfa, self.xrftm, self.ird, self.dstd, n]



    def startpro(self, pdf_file):

        # Specify the path to your Word file
        pdf_file_path = pdf_file
        text_data = self.extract_text_from_pdf(pdf_file_path)

        self.ndata = self.newdata(text_data)

        self.name,self.id, self.manuf, self.cont, self.obs, self.totaldata, self.rpd, self.dap, self.drp, self.ppa, self.psa, self.xfm, self.cfa, self.xrftm, self.ird, self.dstd, self.events = self.ndata

     ###### make all list without words######
        self.name[0] = self.name[0].replace(':',' ').replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.capname = self.name[0].split()

        self.id[0]= self.id[0].replace(':',' ').replace('(', '').replace(')', ' ').replace('*', '').replace(',', '')
        self.name_id = self.id[0].split()

        self.manuf[0] = self.manuf[0].replace(':',' ').replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.manufacturer = self.manuf[0].split()

        self.cont[0] = self.cont[0].replace(':',' ').replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.content = self.cont[0].split()

        self.obs[1] = self.obs[1].replace('(', '').replace(')', '').replace('*', '').replace(',', '')
        self.observer = self.obs[1].split()

        self.total = []

        for item in self.totaldata:
            # Remove 'Gy.m2' from the string
            value = item.replace('Gy.m2', '')

            # Remove 'Gy.m' and 'cm' from the string
            value = value.replace('Gy.m', '').replace('cm', '').replace('s', '')

            # If 'Gy' is still present, remove it
            if 'Gy' in value:
                value = value.replace('Gy', '')

            if value:
                self.total.append(float(value))
            else:
                # If the value is empty, append None to the list
                self.total.append(None)
            # Convert the remaining string to a float and append to the numbers list

        self.moddap = [item.split(' ')[0] for item in self.dap]
        self.dap1 = []
        for item in self.moddap:
            try:
                # Try converting to int first
                num = int(item)
            except ValueError:
                try:
                    # If it's not an int, try converting to float
                    num = float(item)
                except ValueError:
                    # If neither int nor float, set it to None
                    num = None
            self.dap1.append(num)

        self.drp1 = [float(re.search(r'\d+\.\d+e[+-]\d+', value).group()) for value in self.drp]

        modpsa = [item.split(' ')[0] for item in self.psa]
        self.psa1 = []
        for item in modpsa:
            try:
                # Try converting to int first
                num = int(item)
            except ValueError:
                try:
                    # If it's not an int, try converting to float
                    num = float(item)
                except ValueError:
                    # If neither int nor float, set it to None
                    num = None
            self.psa1.append(num)

        self.cfa1 = [float(re.search(r'\d+\.\d+e[+-]\d+', value).group()) for value in self.cfa]


        modxrftm = [item.split(' ')[0] if isinstance(item, str) else item for item in self.xrftm]
        self.xrftm1 = []

        for item in modxrftm:
            if isinstance(item, str):
                try:
                    # Try converting to int first
                    num = int(item)
                except ValueError:
                    try:
                        # If it's not an int, try converting to float
                        num = float(item)
                    except ValueError:
                        # If neither int nor float, set it to None
                        num = None
            else:
                # If the item is not a string, set it to None in the result as well
                num = None

            self.xrftm1.append(num)

        self.ird1 = []
        for item in self.ird:
            numeric_sequences = re.findall(r'\b\d+(?:\.\d+)?\b', item)

            # Add each numeric sequence to the result
            for seq in numeric_sequences:
                self.ird1.append(seq)
        self.ird1 = [float(num) if '.' in num else int(num) for num in self.ird1]
        self.dstd1 = []
        for item in self.dstd:
            numeric_sequences = re.findall(r'\b\d+(?:\.\d+)?\b', item)

            # Add each numeric sequence to the result
            for seq in numeric_sequences:
                self.dstd1.append(seq)
        self.dstd1 = [float(num) if '.' in num else int(num) for num in self.dstd1]


        # Create a pandas DataFrame from the extrac+ted data
        self.all_data = { "Dose Area Product (Gym2)":self.dap1,"Dose (RP) (Gy)":self.drp1,'Collimated Field Area (m2)':self.cfa1,
                          "X-Ray Filter Material":self.xfm,
                           'X-Ray Filter Thickness (mmCu)':self.xft,'Irradiation Duration (s)':self.ird1}


        # Use regular expressions to extract numbers from mixed elements
        max_length = max(len(self.all_data[col]) for col in self.all_data)
        for col in self.all_data:
            self.all_data[col] += [np.nan] * (max_length - len(self.all_data[col]))

        self.df = pd.DataFrame(self.all_data)
        for i in range(0,self.events):
            self.df = self.df.rename(index={i: "Event {0}".format(i+1)})
        self.df = self.df.rename_axis('Irradiation Event X-Ray Data of '+ self.capname[1]+" "+self.name[1])
        #print(self.df)
        self.data_total = {"ID": self.name_id[0],"Manufacturer":self.manufacturer[1],"Content Date":self.content[1] ,"Observer":self.observer[0], "Dose Area Product Total (Gym2)": self.total[0],
                           "Dose (RP) Total (Gy)":self.total[1],"Fluoro Dose Area Product Total (Gym2)":self.total[2],"Fluoro Dose (RP) Total (Gy)":self.total[3],
                           "Total Fluoro Time (s)":self.total[4],"Acquisition Dose Area Product Total (Gym2)":self.total[5],
                            "Acquisition Dose (RP) Total (Gy)":self.total[6],	"Total Acquisition Time (s)":self.total[7],"Reference Point Definition (mm)":self.total[8]}
        self.dft = pd.DataFrame(self.data_total, index=[self.capname[1]+" "+self.name[1]])
        self.dft = self.dft.rename_axis('Accumulated X-Ray Dose Data')

        self.person_data = [self.capname[1]+" "+self.name[1],self.name_id[0],self.manufacturer[1],self.content[1],self.observer[0],self.events]

        self.dfper = pd.DataFrame(self.person_data, index=['Patient name', 'ID', 'Manufacturer', 'Content Date', 'Person Observer Name', 'Number of irradiation events'], columns=[""])


        return self.df, self.dft, self.dfper, self.name_id[0]
        # Output directory where Excel files will be saved


    def get_dataframes(self):
        return self.df, self.dft, self.dfper, self.name_id[0]
        
