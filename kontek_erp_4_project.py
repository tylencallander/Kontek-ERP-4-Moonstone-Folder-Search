import os
import json
import openpyxl
import re

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

def search_project_folders(base_paths, project_numbers, job_numbers):
    found_projects = {}
    errors = {'not_found': [], 'found_in_U_not_in_P': [], 'found_in_both': []}
    projects_in_u = set()

    pattern = re.compile(r"(M\d+)")
    for base_path in base_paths:
        for root, dirs, files in os.walk(base_path):
            for dir in dirs:
                matches = pattern.findall(dir)
                if any(num in project_numbers or num in job_numbers for num in matches):
                    full_path = os.path.join(root, dir)
                    if 'U:' in base_path:
                        projects_in_u.add(dir)
                        if dir in found_projects:
                            errors['found_in_both'].append(dir)
                        else:
                            errors['found_in_U_not_in_P'].append(dir)
                    else:
                        found_projects[dir] = full_path
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
