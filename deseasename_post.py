import json
import time
from datetime import date, datetime
from html import escape
from pathlib import Path

import openpyxl


ORCA_DISEASE_URL = "http://ormaster:ormaster@172.16.123.100:8000/api/orca22/diseasev3"
EXCEL_PATH = Path(__file__).with_name("Dr.入力病名一覧.xlsx")


def format_date(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def format_cell(value):
    if value is None:
        return ""
    return str(value).strip()


def format_outcome(outcome_name, fallback=""):
    outcome_name = format_cell(outcome_name)
    if outcome_name == "治癒":
        return "3"
    if outcome_name == "中止（転医）":
        return "N"
    return format_cell(fallback)


def byoumei_post(
    month,
    ptid,
    disease_name,
    start_date,
    outcome="",
    end_date="",
    department_code="03",
    perform_date=None,
):
    import requests
    import xmltodict

    url = ORCA_DISEASE_URL
    headers = {"Content-Type": "application/xml"}
    perform_date = perform_date or start_date or date.today().strftime("%Y-%m-%d")

    xml_data = f"""
<data>
	<diseasereq type="record">
		<Patient_ID type="string">{escape(str(ptid))}</Patient_ID>
		<Base_Month type="string">{escape(str(month))}</Base_Month>
		<Perform_Date type="string">{escape(str(perform_date))}</Perform_Date>
		<Perform_Time type="string"></Perform_Time>
		<Diagnosis_Information type="record">
			<Department_Code type="string">{escape(str(department_code))}</Department_Code>
		</Diagnosis_Information>
		<Disease_Information type="array">
			<Disease_Information_child type="record">
				<Disease_Insurance_Class type="string"></Disease_Insurance_Class>
				<Disease_Code type="string"></Disease_Code>
				<Disease_Name type="string">{escape(str(disease_name))}</Disease_Name>
				<Disease_Single type="array">
					<Disease_Single_child type="record">
						<Disease_Single_Code type= "string"></Disease_Single_Code>
						<Disease_Single_Name type= "string"></Disease_Single_Name>
					</Disease_Single_child>
				</Disease_Single>
				<Disease_Supplement_Name type= "string"></Disease_Supplement_Name>
				<Disease_Supplement_Single type="array">
					<Disease_Supplement_Single_child type="record">
						<Disease_Supplement_Single_Code type= "string"></Disease_Supplement_Single_Code>
					</Disease_Supplement_Single_child>
				</Disease_Supplement_Single>
				<Disease_InOut type="string"></Disease_InOut>
				<Disease_Category type="string"></Disease_Category>
				<Disease_SuspectedFlag type="string"></Disease_SuspectedFlag>
				<Disease_StartDate type="string">{escape(str(start_date))}</Disease_StartDate>
				<Disease_EndDate type="string">{escape(str(end_date))}</Disease_EndDate>
				<Disease_OutCome type="string">{escape(str(outcome))}</Disease_OutCome>
				<Disease_Karte_Name type="string"></Disease_Karte_Name>
				<Disease_Class type="string">Auto</Disease_Class>
				<Insurance_Combination_Number type="string"></Insurance_Combination_Number>
				<Disease_Receipt_Print type="string"></Disease_Receipt_Print>
				<Disease_Receipt_Print_Period type="string"></Disease_Receipt_Print_Period>
				<Insurance_Disease type="string"></Insurance_Disease>
				<Discharge_Certificate type="string"></Discharge_Certificate>
				<Main_Disease_Class type="string"></Main_Disease_Class>
				<Sub_Disease_Class type="string"></Sub_Disease_Class>
			</Disease_Information_child>
		</Disease_Information>
	</diseasereq>
</data> 
    """

    response = requests.post(url, headers=headers, data=xml_data.encode("utf-8"))
    response.raise_for_status()

    xml_data = response.content.decode("utf-8")
    xml_dict = xmltodict.parse(xml_data)
    return json.dumps(xml_dict, indent=2, ensure_ascii=False)


def read_diseases_from_excel(excel_path=EXCEL_PATH):
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    worksheet = workbook.active
    headers = [format_cell(cell.value) for cell in worksheet[1]]

    for row_number, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
        record = dict(zip(headers, row))
        ptid = format_cell(record.get("ptid"))
        disease_name = format_cell(record.get("byomei"))
        start_date = format_date(record.get("sryymd"))

        if not ptid or not disease_name or not start_date:
            print(f"skip row {row_number}: ptid/byomei/sryymd のいずれかが空です")
            continue

        yield {
            "row_number": row_number,
            "ptid": ptid,
            "disease_name": disease_name,
            "start_date": start_date,
            "outcome": format_outcome(record.get("tenkimei"), record.get("tenkikubun")),
            "end_date": format_date(record.get("tenkiymd")),
        }


def post_diseases_from_excel(excel_path=EXCEL_PATH, interval_seconds=0.3):
    results = []
    for disease in read_diseases_from_excel(excel_path):
        month = disease["start_date"][:7]
        result = byoumei_post(
            month=month,
            ptid=disease["ptid"],
            disease_name=disease["disease_name"],
            start_date=disease["start_date"],
            outcome=disease["outcome"],
            end_date=disease["end_date"],
        )
        print(f"posted row {disease['row_number']}: {disease['ptid']} {disease['disease_name']}")
        results.append({"disease": disease, "result": result})
        time.sleep(interval_seconds)
    return results


if __name__ == "__main__":
    post_diseases_from_excel()