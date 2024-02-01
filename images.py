from PIL import Image
import pytesseract
import pandas as pd
import numpy as np
import re


class MakeExcelFromTiff:

    def extract_text_from_tiff(self, tiff_file_path):
        try:
            with Image.open(tiff_file_path) as image:
                # Use Tesseract to do OCR on the image
                text = pytesseract.image_to_string(image, lang='eng')
                return text

        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def is_date_string(self, s):
        date_pattern = re.compile(r'\d{2}-[a-zA-Z]{3}-\d{2}')
        return bool(date_pattern.match(s))

    def new_data(self, ex_data):

        ex_data = "".join(ex_data)
        ex_data = ex_data.split()
        print(ex_data)

        self.name = []
        self.id = []
        self.manuf = []
        self.cont = []
        self.obs = []
        self.num_exp = []
        self.series = []
        self.kvp = []
        self.current = []
        self.pulwid = []
        self.xrft = []
        self.field = []
        self.kap = []
        self.rp = []
        self.time = []
        self.pulse = []
        self.primang = []
        self.sec = []
        self.num_fr = []
        number_list = [self.convert_text_to_number(item) if isinstance(item, str) else item for item in ex_data]
        # Create a new list with date strings
        self.cont = [self.cont for self.cont in ex_data if self.is_date_string(str(self.cont))]

        self.id.extend([number_list[i] for i in range(len(number_list)) if
                        isinstance(number_list[i], int) and 100000 <= number_list[i] <= 99999999])

        for i in range(0, len(ex_data)):
            if ex_data[i] == "Name:":
                words = (str(ex_data[i + 1]) + str(ex_data[i + 2])).split()
                capital_words = [word for word in words if word.isupper() and not any(char.isdigit() for char in word)]
                if not capital_words:
                    self.name.append("DefaultName")
                else:
                    self.name.append(" ".join(capital_words).replace(",", " "))

            elif ex_data[i] == 'Physician:':
                words = (str(ex_data[i + 1]) + str(ex_data[i + 2])).split()
                capital_words_obs = [word for word in words if
                                     word.isupper() and not any(char.isdigit() for char in word)]
                if not capital_words_obs:
                    self.obs.append("DefaultObs")
                else:
                    self.obs.append(" ".join(capital_words_obs).replace(",", " "))

            if isinstance(ex_data[i], str):
                if ex_data[i] == "CARD":
                    numbers = re.findall(r'\d+', ex_data[i - 1])
                    self.series.extend([int(num) for num in numbers if int(num) < 100])
                elif re.match(r'\d+mA|\d+\.\d+mA', ex_data[i]):
                    self.current.append(ex_data[i])  # tube current
                elif re.match(r'\d+[kK]Y|\d+[kK]|\d+kV|\d+\.\d+[kK]Y|\d+\.\d+[kK]|\d+\.\d+kV', ex_data[i]):
                    self.kvp.append(ex_data[i])  # kv
                elif re.match(r'-\d+pGym\?|\d+nGym\?|\d+Tyeym\*|\d+n@ym\*|\d+pGym\?|\d+\.\d+pGym\?|\d+\.\d+nGym\?|\d+\.\d+Tyeym\*|\d+\.\d+n@ym\*',
                              ex_data[i]):
                    if re.match(r'^\d+\.$', ex_data[i - 1]):
                        self.kap.append(ex_data[i - 1] + ex_data[i])
                    else:
                        self.kap.append(ex_data[i])  # Acquisition Dose Area Product (μGym^2)
                elif re.match(r'\d+\.\d+mGy|\d+mGy|\d+mey|\d+\.+imGy|\d+\.\d+mey', ex_data[i]):
                    if re.match(r'^\d+\.$', ex_data[i - 1]):
                        self.rp.append(ex_data[i - 1] + ex_data[i])
                    else:
                        self.rp.append(ex_data[i])  # Acquisition Dose (RP) (Gy)
                elif re.match(r'(0\.0cu|\d+cu|\d+Cu|\d+Ccu|\d+\.\d+cu|\d+\.\d+Cu)', ex_data[i]):
                    self.xrft.append(ex_data[i])  # X-ray Filter Thickness (Cu)
                elif re.match(r'\d+cm|\d+\.\d+cm', ex_data[i]):
                    self.field.append(ex_data[i])  # Field Width (cm)
                elif re.match(r'(\d+LAO|\d+LA0|\d+RAO)', ex_data[i]):
                    self.primang.append(ex_data[i])  # Primary Angle (deg)  # 'LAO', 'LA0', 'RAO'
                elif re.match(r'\d+CRA', ex_data[i]):
                    self.sec.append(ex_data[i])  # Secondary Angle (deg)
                elif re.match(r'\d+\.\d+ms|\d+ms', ex_data[i]):
                    if re.match(r'^\d+\.$', ex_data[i - 1]):
                        self.pulwid.append(ex_data[i - 1] + ex_data[i])
                    else:
                        self.pulwid.append(ex_data[i])  # Pulse Width (ms)

                elif re.match(r'\d+s|\d+\.\d+s', ex_data[i]):
                    self.time.append(ex_data[i])  # Acquisition Time (s)
                elif re.match(r'\d+F/s|\d+\.\d+F/s', ex_data[i]):
                    self.pulse.append(ex_data[i])  # Frame Rate (pulse/s)
                elif re.match(r'\d+F', ex_data[i]):
                    self.num_fr.append(ex_data[i])  # Number Frames
                elif ex_data[i] == 'Exposures:':
                    self.num_exp.append(ex_data[i + 1])  # number of exposures

        self.kvp = [
            str(item).replace("k", "").replace("kY", "").replace("Y", "").replace("V", "").replace("y", "").replace("v", "")
            if isinstance(item, (str, int, float)) else item for item in self.kvp]
        self.series = [str(item).replace("F", "") if isinstance(item, (str, int, float)) else item for item in
                       self.series]
        self.current = [
            str(item).translate(str.maketrans('', '', 'mAkyYV')) if isinstance(item, (str, int, float)) else item for
            item in self.current]

        self.pulwid = [re.sub(r'\.\d+CL|ms|_|-', '', str(item)) if isinstance(item, (str, int, float)) else item for item in self.pulwid]

        self.xrft = [str(item).replace("cu", "").replace("C", "").replace("Cu", "").replace("u", "") if isinstance(item, (str, int, float)) else item for
                     item in self.xrft]
        self.field = [str(item).replace("cm", "") if isinstance(item, (str, int, float)) else item for item in
                      self.field]
        self.kap = [
            str(item).replace("pGym?", "").replace("yGym?", "").replace("nGym?", "").replace("Tyeym*", "").replace(
                "n@ym*", "") if isinstance(item, (str, int, float)) else item for item in self.kap]
        self.rp = [str(item).replace("mGy", "").replace("mey", "").replace("\d.nGym?", "") if isinstance(item, (
        str, int, float)) else item for item in self.rp]
        self.primang = [str(item).replace("LAO", "").replace("LA0", "").replace("RAO", "") if isinstance(item, (
        str, int, float)) else item for item in self.primang]
        self.sec = [str(item).replace("CRA", "") if isinstance(item, (str, int, float)) else item for item in self.sec]
        self.num_fr = [str(item).replace("F", "") if isinstance(item, (str, int, float)) else item for item in
                       self.num_fr]
        try:
            self.current = [float(element) for element in self.current]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass
        try:
            self.pulwid = [float(element) for element in self.pulwid]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass
        try:
            self.rp = [float(element) for element in self.rp]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        try:
            self.kap = [float(element) for element in self.kap]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass
        try:
            self.xrft = [float(element) for element in self.xrft]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass
        try:
            self.field = [float(element) for element in self.field]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass
        '''self.primang = [int(element) for element in self.primang]
        self.sec = [float(element) for element in self.sec]
        self.num_fr = [float(element) for element in self.num_fr]'''

        self.time = [str(item).replace("s", "") if isinstance(item, (str, int, float)) else item for item in self.time]
        self.pulse = [str(item).replace("F/s", "") if isinstance(item, (str, int, float)) else item for item in
                      self.pulse]
        try:
            self.time = [float(element) for element in self.time]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass
        try:
            self.pulse = [float(element) for element in self.pulse]
        except ValueError as e:
            print(f"Error converting to float: {e}")
        pass

        return [self.name, self.id, self.cont, self.obs, self.num_exp, self.series, self.kvp, self.current, self.pulwid,
                self.xrft, self.field, self.time, self.pulse, self.kap, self.rp, self.primang, self.sec, self.num_fr]

    def start_processing(self, tiff_file):
        # Use the new Tiff extraction method
        text_data = self.extract_text_from_tiff(tiff_file)

        self.ndata = self.new_data(text_data)

        self.name, self.id, self.cont, self.obs, self.num_exp, self.series, self.kvp, self.current, self.pulwid, self.xrft, self.field, self.time, self.pulse, self.kap, self.rp, self.primang, self.sec, self.num_fr = self.ndata

        # self.kvp = [item.replace("kV", "") for item in self.kvp]

        self.data_total = {"Exposure Data": self.series, "KVP": self.kvp, "Tube Current (mA)": self.current,
                           "Pulse Width (ms)": self.pulwid, "X-ray Filter Thickness (Cu)": self.xrft,
                           "Field Width (cm)": self.field, "Acquisition Time (s)": self.time,
                           "Frame Rate (pulse/s)": self.pulse, "Acquisition Dose Area Product (μGym^2)": self.kap,
                           "Acquisition Dose (RP) (Gy)": self.rp, "Primary Angle (deg)": self.primang,
                           "Secondary Angle (deg)": self.sec, "Number Frames": self.num_fr}

        # Find the maximum length among all lists
        max_length = max(len(v) for v in self.data_total.values())

        # Auto insert np.nan for missing values
        for key, value in self.data_total.items():
            self.data_total[key] = value + [np.nan] * (max_length - len(value))

        print(self.data_total)
        min_length = min((len(self.data_total[key]) for key in self.data_total))

        # Check which lists have a length different from the minimum length
        for key, values in self.data_total.items():
            if len(values) != min_length:
                print(
                    f"The list '{key}' has a different length ({len(values)}) than the minimum length ({min_length}).")

        self.dft = pd.DataFrame(self.data_total)
        self.dft['Frame Rate (pulse/s)'] = pd.to_numeric(self.dft['Frame Rate (pulse/s)'], errors='coerce')
        self.dft = self.dft.set_index("Exposure Data")
        # Ensure each list has at least one element
        self.name = self.name + ['DefaultName'] if not self.name else self.name
        self.id = self.id + ['DefaultID'] if not self.id else self.id
        self.cont = self.cont + ['DefaultCont'] if not self.cont else self.cont
        self.obs = self.obs + ['DefaultObs'] if not self.obs else self.obs
        self.num_exp = self.num_exp + ['DefaultNumExp'] if not self.num_exp else self.num_exp

        self.person_data = [self.name[0], self.id[0], self.cont[0], self.obs[0], self.num_exp[0]]

        self.dfper = pd.DataFrame(self.person_data,
                                  index=['Patient name', 'ID', 'Date', 'Performing physician', 'Number of exposures'],
                                  columns=[""])

        return self.dft, self.dfper

    def convert_text_to_number(self, text):
        try:
            # Extract numeric part
            numeric_part = ''.join(filter(lambda x: x.isdigit() or x == '.' or x == '-', text))

            if numeric_part:
                # Try converting to different numeric types
                try:
                    return int(numeric_part)
                except ValueError:
                    return float(numeric_part)
            else:
                return text  # Return the original text if no numeric part is found

        except ValueError:
            # Handle the case where conversion is not possible
            return text  # Return the original text if an error occurs
