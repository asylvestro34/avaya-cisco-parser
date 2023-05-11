from contextlib import contextmanager
from collections import namedtuple
from copy import deepcopy
import csv
import re
import os
import time

from jinja2 import Template
import pandas as pd


@contextmanager
def ignored(*exceptions):
    try:
        yield
    except exceptions:
        pass


ButtonMapResult = namedtuple("ButtonMapResult", ["cisco_button_function", "button_type"])


def get_abrv_dial_mapping():
    get_abrv_dial_mapping = pd.read_excel("abrv_data/abrv-dial.xlsx", header=0, na_filter=False).values
    return get_abrv_dial_mapping


class Phone:
    abrv_dial_mapping = get_abrv_dial_mapping()

    """
    # Uncomment this if dummy numbers are to be used as a proxy for this type of button
    aut_msg_wait_base = "130185"
    aut_msg_wait_start = 1000
    """

    def __init__(
        self,
        extension,
        fullname,
        firstname,
        lastname,
        type="",
        port="",
        coverage_path_1="",
        coverage_path_2="",
        cor="",
        cos="",
        ec500="",
        ip_softphone="",
        system_number="",
        has_expansion="",
    ):
        self.extension = extension
        self.fullname = fullname
        self.firstname = firstname
        self.lastname = lastname
        self.type = type
        self.port = port
        self.coverage_path_1 = coverage_path_1
        self.coverage_path_2 = coverage_path_2
        self.cor = cor
        self.cos = cos
        self.ec500 = ec500
        self.ip_softphone = ip_softphone
        self.system_number = system_number
        self.has_expansion = has_expansion

        self.buttons = list()

    def add_buttons(self, button_parse, mod1=False, mod2=False):
        buttons = re.findall(r"(\d+:.*?(\n|(?=\d+:)|$))", button_parse)

        for button in buttons:
            button_num = ""
            button_type = ""
            button_extension = ""
            button_label = ""
            button_pair = None
            button_other = ""
            button_abrv_list_num = ""
            button_abrv_list_name = ""
            button_abrv_dc = ""

            with ignored(AttributeError):
                button_num = re.search(r"(\d+):", button[0]).group(1)
            if len(button_num) == 1:
                button_num = "0" + button_num
            if mod1 == True:
                button_num = str(int(button_num) + 24)
            elif mod2 == True:
                button_num = str(int(button_num) + 48)
            with ignored(AttributeError):
                button_type = re.search(r"\d+:\s*(.*?)\s+", button[0]).group(1)
            if button_num == "01":
                button_extension = self.extension
            else:
                with ignored(AttributeError):
                    button_extension = re.search(r":\s*([a-z\-0-9]+)(\s{1}).*?([0-9\-]{4,20})", button[0]).group(3)
                if not button_extension:
                    button_extension = "No Ext"
                else:
                    button_extension = button_extension.replace("-", "")

            if button_type:
                with ignored(AttributeError):
                    button_other = re.search(r"(R|Rg):(\w+)", button[0]).group(2)
                if button_type == "abrv-dial":
                    with ignored(AttributeError):
                        abbreviated_dialing = re.findall(
                            r"\nABBREVIATED\sDIALING\s+\n\s+List1:\s(.+?)\s+List2:\s(.+?)\s+List3:\s(.+?)\s+\n",
                            seperated,
                        )
                    abrv_dial_lists = abbreviated_dialing[0]

                    abrv_dial_extra_info = re.search(r"abrv-dial\s*List:\s*(\d+)\s*DC:\s*(\d+)", button[0])
                    abrv_dial_list_number = int(abrv_dial_extra_info.group(1))
                    button_abrv_list_num = abrv_dial_list_number
                    abrv_dial_dc = int(abrv_dial_extra_info.group(2))
                    button_abrv_dc = abrv_dial_dc
                    abrv_dial_mappings = Phone.abrv_dial_mapping
                    button_abrv_list_name = abrv_dial_lists[abrv_dial_list_number-1]
                    for mapping in abrv_dial_mappings:
                        if button_abrv_list_name == mapping[0]:
                            if abrv_dial_dc == mapping[1]:
                                button_extension = mapping[2]
                                button_label = mapping[3]
                                if button_label == "":
                                    button_label = button_extension
                    if not button_extension:
                        continue
                elif button_type == "aut-msg-wt":
                    button_label = button_extension
                    """
                    # Uncomment this if dummy numbers are to be used as a proxy for this type of button
                    button_extension = Phone.aut_msg_wait_base + str(
                        Phone.aut_msg_wait_start
                    )
                    Phone.aut_msg_wait_start += 1
                    """
                elif button_type == "call-appr":
                    button_label = button_extension
                elif button_type == "hunt-ns":
                    with ignored(AttributeError):
                        button_other = re.search(r"Grp:\s*(\d+)", button[0]).group(1)
                elif button_type == "dial-icom":
                    with ignored(AttributeError):
                        button_extension = "Grp: " + re.search(r"Grp:\s*(\d+)", button[0]).group(1)
                elif button_type == "autodial":
                    with ignored(AttributeError):
                        button_extension = re.search(r":\s*([a-z\-0-9]+)\s+Number:\s+(.*?)(\s+|$)", button[0]).group(2)
                    if not button_extension:
                        button_extension = "No Ext"
                    else:
                        button_extension = button_extension.replace("-", "")

                button_pair = (button_type, button_extension)
                with ignored(KeyError):
                    if button_label:
                        self.buttons.append(Button.from_pair(button_pair, button_label, button_other=button_other, button_num=button_num, button_abrv_list_num=button_abrv_list_num, button_abrv_list_name=button_abrv_list_name, button_abrv_dc=button_abrv_dc))
                    else:
                        self.buttons.append(Button.from_pair(button_pair, button_other=button_other, button_num=button_num, button_abrv_list_num=button_abrv_list_num, button_abrv_list_name=button_abrv_list_name, button_abrv_dc=button_abrv_dc))

    @property
    def row_dict(self):
        return dict(
            extension=self.extension,
            fullname=self.fullname,
            firstname=self.firstname,
            lastname=self.lastname,
            port=self.port,
            type=self.type,
            coverage_path_1=self.coverage_path_1,
            coverage_path_2=self.coverage_path_2,
            cor=self.cor,
            cos=self.cos,
            ec500=self.ec500,
            ip_softphone=self.ip_softphone,
            has_expansion=self.has_expansion
        )

    @property
    def buttons_plk(self):
        return [button for button in self.buttons if button.button_type.upper() == 'PLK']

    @property
    def plk_count(self):
        return len([button for button in self.buttons if button.button_type.upper() == 'PLK'])

    @property
    def line_count(self):
        return len([button for button in self.buttons if 'line' in button.cisco_button_function.lower()])

    @property
    def speed_count(self):
        return len([button for button in self.buttons if 'speed' in button.cisco_button_function.lower()])

    @property
    def has_blf(self):
        return any([button for button in self.buttons if 'BLF' in button.cisco_button_function])

    def __key(self):
        return (self.extension, self.fullname, self.type, self.port)

    def __hash__(self):
        return hash(self.__key())

    def __eq__(self, other):
        return isinstance(self, type(other)) and self.__key() == other.__key()


class Button:
    mapping = {
        "abrdg-appr": ButtonMapResult("Shared-line", "PLK"),
        "abrv-dial": ButtonMapResult("Speed-dial", "PLK"),
        "aut-msg-wt": ButtonMapResult("Line (ring)", "PLK"),
        "auto-cback": ButtonMapResult("Redial", "Soft-Key"),
        "auto-icom": ButtonMapResult("Intercom", "PLK"),
        "autodial": ButtonMapResult("Speed-dial", "PLK"),
        "brdg-appr": ButtonMapResult("Shared-line", "PLK"),
        "busy-ind": ButtonMapResult("BLF", "PLK"),
        "call-appr": ButtonMapResult("Line (ring)", "PLK"),
        "call-fwd": ButtonMapResult("Forward all", "Soft-Key"),
        "call-park": ButtonMapResult("Call park", "Soft-Key"),
        "call-pkup": ButtonMapResult("Pickup", "Soft-Key"),
        "conf-dsp": ButtonMapResult("Conference", "Soft-Key"),
        "dial-icom": ButtonMapResult("Intercom", "PLK"),
        "directory": ButtonMapResult("Directory", "Physical Key"),
        "dn-dst": ButtonMapResult("Do not Disturb", "Soft-Key"),
        "drop": ButtonMapResult("End Call", "Physical Key"),
        "ec500": ButtonMapResult("Mobility", "Soft-Key"),
        "extnd-call": ButtonMapResult("Mobility", "Soft-Key"),
        "goto-cover": ButtonMapResult("Decline", "Soft-Key"),
        "grp-page": ButtonMapResult("Speed-dial", "PLK"),
        "headset": ButtonMapResult("Headset", "Physical Key"),
        "hunt-ns": ButtonMapResult("Hunt group logout", "PLK"),
        "inst-trans": ButtonMapResult("Speed-dial", "PLK"),
        "mct-act": ButtonMapResult("Malicious Call Trace", "Soft-Key"),
        "no-hld-cnf": ButtonMapResult("Conference", "Physical Key"),
        "release": ButtonMapResult("End Call", "Physical Key"),
        "send-calls": ButtonMapResult("Do not Disturb", "Soft-Key"),
        "serv-obsrv": ButtonMapResult("serv-obsrv", "PLK"),
    }

    def __init__(
        self,
        avaya_button_function,
        cisco_button_function="",
        button_type="",
        extension="",
        label="",
        button_other="",
        button_num="",
        button_abrv_list_num="",
        button_abrv_list_name="",
        button_abrv_dc="",
    ):
        self.avaya_button_function = avaya_button_function
        self.cisco_button_function = cisco_button_function
        self.button_type = button_type
        self.extension = extension
        self.label = label
        self.button_other = button_other
        self.button_num = button_num
        self.button_abrv_list_num = button_abrv_list_num
        self.button_abrv_list_name = button_abrv_list_name
        self.button_abrv_dc = button_abrv_dc

    @classmethod
    def from_pair(cls, pair, label=None, button_other="", button_num="", button_abrv_list_num="", button_abrv_list_name="", button_abrv_dc=""):
        label = label if label else pair[1]
        return cls(
            pair[0],
            Button.mapping[pair[0]].cisco_button_function,
            Button.mapping[pair[0]].button_type,
            pair[1],
            label,
            button_other,
            button_num,
            button_abrv_list_num,
            button_abrv_list_name,
            button_abrv_dc,
        )

    def __key(self):
        return (self.avaya_button_function, self.extension)

    def __hash__(self):
        return hash(self.__key())

    def __eq__(self, other):
        return isinstance(self, type(other)) and self.__key() == other.__key()


def return_file_list(substring):
    return [file for file in collaboration_export_files if substring in file.lower()]


def get_system_number(matchobject):
    system_number = 0

    with ignored(AttributeError):
        if matchobject:
            for index, match in enumerate(matchobject[0], 1):
                if match.lower() == "system":
                    system_number = index
        return int(system_number)


def write_excel(rows, file, reorder=None, sort=None):
    df = pd.DataFrame(rows)


    if not reorder:
        columns = df.columns.values.tolist()

        columns_with_numbers = [
            column for column in columns if re.search(r"\d+", column)
        ]
        columns_without_numbers = [
            column for column in columns if not re.search(r"\d+", column)
        ]

        columns_with_numbers.sort(key=lambda x: int("".join(filter(str.isdigit, x))))

        columns_new = columns_without_numbers + columns_with_numbers

        df = df[columns_new]
    else:
        df = df[reorder]

    if sort:
        df = df.sort_values(sort)

    file_name_no_ext = file.split(".")[0]
    timestr = time.strftime("%Y%m%d-%H%M%S")

    print(f"File name: {file_name_no_ext}-{timestr}.csv")
    df.to_csv(f"{file_name_no_ext}-{timestr}.csv", index=None)


def output_stacked(phone_set):
    rows = []

    for phone in phone_set:
        plk_count = 0
        non_speed_dial_plk = 0
        non_speed_rollovers_plk = 0

        button_list = []
        for button in phone.buttons:
            button_list.append(button.__dict__)
        button_d_list = deepcopy(button_list)
        [d.pop("button_num", None) for d in button_d_list]
        for button in button_d_list[:]:
            if button["extension"] == "No Ext":
                    button_d_list.remove(button)
        seen = set()
        unique_buttons = []
        for dict in button_d_list:
            tup = tuple(dict.items())
            if tup not in seen:
                seen.add(tup)
                unique_buttons.append(dict)
        button_d_list = unique_buttons
        
        for button in button_d_list:
            if button["button_type"] == "PLK":
                plk_count += 1
                if button["cisco_button_function"] != "Speed-dial":
                    non_speed_dial_plk += 1
                    if "line" in button["cisco_button_function"].lower():
                        non_speed_rollovers_plk += 1

        if plk_count == 0:
            plk_count = 1
            non_speed_dial_plk = 1
            non_speed_rollovers_plk = 1

        row_dictionary = phone.row_dict

        row_dictionary_w_counts = deepcopy(row_dictionary)

        row_dictionary_w_counts["plk_count"] = plk_count
        row_dictionary_w_counts["plk_minus_speed_count"] = non_speed_dial_plk
        row_dictionary_w_counts["plk_minus_speed_rollovers_count"] = non_speed_rollovers_plk

        rows.append(row_dictionary_w_counts)

        for button in phone.buttons:
            if button.button_type == "PLK":
                next_row = deepcopy(row_dictionary)
                next_row["button_function_avaya"] = button.avaya_button_function
                next_row["button_function_cisco"] = button.cisco_button_function
                next_row["button_other"] = button.button_other
                next_row["button_num"] = button.button_num
                next_row["button_extension"] = button.extension
                next_row["button_label"] = button.label
                next_row["button_abrv_list_num"] = button.button_abrv_list_num
                next_row["button_abrv_list_name"] = button.button_abrv_list_name
                next_row["button_abrv_dc"] = button.button_abrv_dc
                rows.append(next_row)

    order = [
        "extension",
        "fullname",
        "firstname",
        "lastname",
        "port",
        "type",
        "cos",
        "cor",
        "coverage_path_1",
        "coverage_path_2",
        "ec500",
        "ip_softphone",
        #"button_num",
        "button_function_cisco",
        "button_function_avaya",
        "button_abrv_list_num",
        "button_abrv_list_name",
        "button_abrv_dc",
        "button_extension",
        "button_label",
        "button_other",
        "has_expansion",
        "plk_count",
        "plk_minus_speed_count",
        "plk_minus_speed_rollovers_count",
    ]

    sort = [
        "extension",
        "plk_count",
        #"button_num",
    ]

    #This line removes the "button_num" column from each row.
    [d.pop("button_num", None) for d in rows]

    for row in rows[:]:
        if "button_extension" in row.keys():
            if row["button_extension"] == "No Ext":
                if row["button_function_avaya"] == "serv-obsrv":
                    continue
                else:
                    rows.remove(row)

    seen = set()
    unique_rows = []
    for dict in rows:
        tup = tuple(dict.items())
        if tup not in seen:
            seen.add(tup)
            unique_rows.append(dict)

    rows = unique_rows

    write_excel(rows, "output-avaya-vec-html.csv", order, sort)


def output_dsr_import(phone_set):
    rows = []

    for phone in phone_set:
        #Manually update the site code.
        site = 'Ross'

        namecheck = phone.fullname.split("_ ")

        if len(namecheck) == 2:
            row_dictionary = dict(seven_digit_extension=phone.extension,
                                site=site,
                                firstname=phone.fullname.split("_ ")[1],
                                lastname=phone.fullname.split("_ ")[0])
        else:
            row_dictionary = dict(seven_digit_extension=phone.extension,
                                site=site,
                                firstname=phone.fullname)

        if phone.extension == '80029':
            for button in phone.buttons:
                print(button.avaya_button_function, button.extension)

        filtered_buttons = [button for button in phone.buttons if 'appr' in button.avaya_button_function.lower() and len(button.extension) >= 4]

        row_dictionary['b1_seven_digit_extension'] = phone.extension
        row_dictionary['b1_button_label'] = phone.extension
        row_dictionary['b1_button_type'] = 'Line (ring)'

        for index, button in enumerate(filtered_buttons, 2):
            dsr_button_type = button.cisco_button_function
            if 'line' in button.cisco_button_function.lower():
                dsr_button_type = 'Line (ring)'

            if index <= 5:
                row_dictionary[f'b{index}_seven_digit_extension'] = button.extension
                row_dictionary[f'b{index}_button_label'] = button.label
                row_dictionary[f'b{index}_button_type'] = dsr_button_type
            else:
                reset_index = index
                row_dictionary[f'b{reset_index}_kem_extension'] = button.extension
                row_dictionary[f'b{reset_index}_kem_label'] = button.label
                row_dictionary[f'b{reset_index}_kem_type'] = dsr_button_type

        rows.append(row_dictionary)

    write_excel(rows, "dsr-import-avaya.csv", None)


def output_bat_import(phone_set):
    rows = []

    with open("data/bat_format.csv", "r", encoding='utf-8-sig') as file:
        data = file.readline().rstrip('\n')
        headers = data.split(',')
        template = Template(data)

    for phone in phone_set:
        rendered = template.render(phone=phone)

        row = dict(zip(headers, rendered.split(',')))
        for i in row:
            if i in phone.row_dict:
                row[i] = phone.row_dict[str(i)]
            else:
                row[i] = ""

        speed_dials = [button for button in phone.buttons if 'speed' in button.cisco_button_function.lower()]

        blf_dials = [button for button in phone.buttons if 'BLF' in button.cisco_button_function]

        shared_lines = [button for button in phone.buttons if 'shared' in button.cisco_button_function.lower()]

        for index, button in enumerate(speed_dials, 1):
            row[f'Speed Dial Number {index}'] = button.extension
            row[f'Speed Dial Label {index}'] = button.extension

        for index, button in enumerate(blf_dials, 1):
            row[f'Busy Lamp Field Destination {index}'] = button.extension
            row[f'Busy Lamp Field Directory Number {index}'] = ''
            row[f'Busy Lamp Field Label {index}'] = button.extension
            row[f'Busy Lamp Field Call Pickup {index}'] = 'f'

        for index, button in enumerate(shared_lines, 2):
            row[f'Directory Number {index}'] = button.extension
            row[f'Route Partition {index}'] = 'PhoneDN-PT'

        rows.append(row)

    write_excel(rows, "bat-import-avaya.csv")


if __name__ == "__main__":
    collaboration_export_folder_name = "station_data"

    collaboration_export_files = os.listdir(collaboration_export_folder_name)

    display_station_files = return_file_list(".")

    print(display_station_files)

    phone_set = set()

    regex_to_search = {
        "fullname": r"Name:\s*(.*?)\s+Coverage",
        "port": r"Port:\s*(.*?)\s+Coverage",
        "type": r"Type:\s*(.*?)\s+Sec",
        "coverage_path_1": r"Coverage\s*Path\s*1:\s*(.*?)\s+COR",
        "coverage_path_2": r"Coverage\s*Path\s*2:\s*(.*?)\s+COS",
        "cor": r"COR:\s*(\d+)",
        "cos": r"COS:\s*(\d+)",
        "ec500": r"EC500\s*State:\s*(\w+)",
        "ip_softphone": r"IP\sSoftPhone\?\s(y|n)",
    }

    for key, value in regex_to_search.items():
        regex_to_search[key] = re.compile(value)

    for file in display_station_files:

        if '.html' in file:
            seperator_regex = r'<H4>Station\s+<.+\n<PRE>'
        elif '.vec' in file:
            seperator_regex = r'Station\s+\d+\s+Details'
        elif '.txt' in file:
            seperator_regex = r'STATION\s*(?=Extension)'
        else:
            continue

        with open(f"{collaboration_export_folder_name}/{file}") as display_station_file:
            display_station_data = display_station_file.read()
            display_station_chunk = re.split(
                seperator_regex, display_station_data
            )[1:]

            for seperated in display_station_chunk:

                if '.html' in file:
                    seperated = re.sub(r'</PRE>\n<HR>\n<B>', '', seperated)

                phone_arguments = dict()
                button_parse = ""
                feature_button_parse = ""
                expansion_button_parse = ""

                phone_arguments["extension"] = (
                    re.search(r"Extension:\s+(.*?)\s+", seperated).group(1).replace("-", "")
                )

                for key, regex in regex_to_search.items():
                    with ignored(AttributeError):
                        phone_arguments[key] = re.search(regex, seperated).group(1).replace(",", "_")

                with ignored(AttributeError):
                    abbreviated_dialing = re.findall(
                        r"\nABBREVIATED\sDIALING\s+\n\s+List1:\s(.+?)\s+List2:\s(.+?)\s+List3:\s(.+?)\s+\n",
                        seperated,
                    )
                with ignored(AttributeError):
                    phone_arguments["system_number"] = get_system_number(
                        abbreviated_dialing
                    )
                with ignored(AttributeError):
                    button_parses = re.findall(
                        r"\nBUTTON ASSIGNMENTS\s*(.*?)(?=[A-Z]{5,25}|$)",
                        seperated,
                        re.DOTALL,
                    )
                with ignored(AttributeError):
                    feature_button_parse = re.findall(
                        r"\nFEATURE BUTTON ASSIGNMENTS\s*(.*?)(?=[A-Z]{5,25}|$)",
                        seperated,
                        re.DOTALL,
                    ).group(1)
                with ignored(AttributeError):
                    expansion_button_parse = re.findall(
                        r"\sEXPANSION MODULE BUTTON ASSIGNMENTS\s*(.*?)(?=[A-Z]{5,25}|$)",
                        seperated,
                        re.DOTALL,
                    ).group(1)
                with ignored(AttributeError):
                    button_modules = re.findall(
                        r"\sBUTTON MODULE #\d+ ASSIGNMENTS\s*(.*?)(?=[A-Z]{5,25}|$)",
                        seperated,
                        re.DOTALL,
                    )

                phone_arguments["has_expansion"] = True if expansion_button_parse else False

                if button_modules:
                    phone_arguments["has_expansion"] = True

                namecheck = phone_arguments["fullname"].split("_ ")
                if len(namecheck) == 2:
                    phone_arguments["firstname"]=phone_arguments["fullname"].split("_ ")[1]
                    phone_arguments["lastname"]=phone_arguments["fullname"].split("_ ")[0]
                else:
                    phone_arguments["firstname"]=phone_arguments["fullname"]
                    phone_arguments["lastname"]=""

                phone = Phone(**phone_arguments)

                for button_parse in button_parses:
                    phone.add_buttons(button_parse)

                phone.add_buttons(feature_button_parse)
                phone.add_buttons(expansion_button_parse)
                
                button_mod_count = 1
                for button_module in button_modules:
                    if button_mod_count == 1:
                        phone.add_buttons(button_module, True, False)
                        button_mod_count += 1
                    if button_mod_count == 2:
                        phone.add_buttons(button_module, False, True)
                phone_set.add(phone)

    output_stacked(phone_set)
    output_dsr_import(phone_set)
    # Don't use the bat_import - it needs to be fixed
    # output_bat_import(phone_set)
