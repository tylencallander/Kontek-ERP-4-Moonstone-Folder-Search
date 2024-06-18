# This ERP is tricky ill come back to it, but the projects.json output is perfect, its just the error sorting.

import os
import json
import openpyxl
import re

# Extracts the project, and serial numbers from the excel file

def extract_project_numbers_from_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active
        project_numbers = set()
        job_numbers = set()
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            serial = str(row[0]).strip().upper() if row[0] else ''
            job = str(row[1]).strip().upper() if row[1] else ''
            if serial.startswith('M') or serial.isdigit():
                project_numbers.add(serial)
            if job.startswith('M'):
                job_numbers.add(job)
            print(f"Extracted project number: {serial} and job number: {job} from Excel")
        return project_numbers, job_numbers
    except Exception as e:
        print(f"Error reading from Excel: {e}")
        return set(), set()

# Searches the P and U base paths for the project folders

def search_project_folders(base_paths, project_numbers, job_numbers):
    found_projects = {}
    errors = {'not_found': [], 'found_in_U_not_in_P': [], 'found_in_both': []}
    projects_in_u = set()

    pattern = re.compile(r"M\d{7}")
    for base_path in base_paths:
        for root, dirs, files in os.walk(base_path):
            for dir in dirs:
                matches = pattern.search(dir)
                if matches:
                    project_key = matches.group()
                    full_path = os.path.join(root, dir)
                    if 'U:' in base_path:
                        projects_in_u.add(project_key)
                        if project_key not in found_projects:
                            errors['found_in_U_not_in_P'].append(project_key)
                        else:
                            errors['found_in_both'].append(project_key)
                    else:
                        if project_key not in projects_in_u:
                            found_projects[project_key] = {
                                "projectnumber": project_key,
                                "projectfullpath": full_path,
                                "projectpath": full_path.split("\\")
                            }
                    print(f"Found project {dir} in {base_path}: {full_path}")

    for num in project_numbers.union(job_numbers):
        if num not in found_projects and num not in projects_in_u:
            errors['not_found'].append(num)

    return found_projects, errors

def main():
    excel_file_path = "P:/MOONSTONE/SOLD MOONSTONE UNITS.xlsx"
    base_paths = [
        "P:/Moonstone/Customer",
        "U:/MOONSTONE/MS Completed Machine Sales",
        "U:/MOONSTONE/MS Non Machine Sales",
        "U:/MOONSTONE/MS Pending Machine Sales"
    ]
    project_numbers, job_numbers = extract_project_numbers_from_excel(excel_file_path)
    projects, errors = search_project_folders(base_paths, project_numbers, job_numbers)

    with open("projects.json", "w") as f:
        json.dump(projects, f, indent=4)
    with open("errors.json", "w") as f:
        json.dump(errors, f, indent=4)

    print("\nParsing Complete!\n")
    print(f"Logged {len(projects)} projects to projects.json")
    print(f"Logged errors to errors.json")

if __name__ == "__main__":
    main()
