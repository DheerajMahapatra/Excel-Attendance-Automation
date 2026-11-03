import io
import random
import re
from datetime import date, datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF00")
NO_FILL = PatternFill(fill_type=None)

ENROLLMENT_PATTERN = re.compile(r"AJU\s*/\s*\d+", re.IGNORECASE)
DATE_HEADER_PATTERN = re.compile(r"^\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}")


def norm_enrollment(value):
    if value is None:
        return None
    match = ENROLLMENT_PATTERN.search(str(value).strip().upper().replace(" ", ""))
    return match.group(0) if match else None


def to_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def find_header_row(ws, required_texts):
    lower_texts = [text.lower() for text in required_texts]
    for row in range(1, min(ws.max_row, 40) + 1):
        values = [
            str(ws.cell(row, col).value or "").strip().lower()
            for col in range(1, ws.max_column + 1)
        ]
        if all(
            any(text in value for value in values)
            for text in lower_texts
        ):
            return row
    raise ValueError(
        "Could not find a header row containing "
        + ", ".join(repr(text) for text in required_texts)
        + f" in sheet '{ws.title}'."
    )


def find_column(headers, patterns):
    for idx, value in enumerate(headers, start=1):
        text = str(value or "").strip().lower()
        if any(pattern in text for pattern in patterns):
            return idx
    return None


def find_total_lectures_column(headers):
    for idx, value in enumerate(headers, start=1):
        text = str(value or "").strip().lower()
        if not text or "attend" in text or "percent" in text or "%" in text:
            continue
        if ("total" in text and ("lecture" in text or "class" in text)) or "conducted" in text:
            return idx
    return None


def is_date_header(value):
    if isinstance(value, (datetime, date)):
        return True
    if isinstance(value, str):
        return bool(DATE_HEADER_PATTERN.match(value))
    return False


def sheet_has_register_layout(ws):
    for row in range(1, min(ws.max_row, 20) + 1):
        values = [
            str(ws.cell(row, col).value or "").strip().lower()
            for col in range(1, ws.max_column + 1)
        ]
        has_register_column = any(
            "regno" in value or "rollno" in value for value in values
        )
        has_date_column = any(
            is_date_header(ws.cell(row, col).value)
            for col in range(1, ws.max_column + 1)
        )
        if has_register_column and has_date_column:
            return True
    return False


def locate_summary_columns(ws):
    enroll_col = None
    attended_col = None
    header_rows = []
    last_header_row = min(ws.max_row, 40)

    for row in range(1, last_header_row + 1):
        for col in range(1, ws.max_column + 1):
            text = str(ws.cell(row, col).value or "").strip().lower()
            if not text:
                continue
            if enroll_col is None and "enrol" in text:
                enroll_col = col
                header_rows.append(row)
            if (
                attended_col is None
                and "attend" in text
                and "percent" not in text
                and "%" not in text
            ):
                attended_col = col
                header_rows.append(row)

    data_start = max(header_rows) + 1 if header_rows else 1

    if enroll_col is None:
        counts = {}
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                if norm_enrollment(ws.cell(row, col).value):
                    counts[col] = counts.get(col, 0) + 1
        if counts:
            enroll_col = max(counts, key=counts.get)
            data_start = 1

    if enroll_col is None:
        raise ValueError(
            f"Could not find an Enrollment No column in sheet '{ws.title}'."
        )

    def numeric_ratio(col):
        total = 0
        valid = 0
        for row in range(data_start, ws.max_row + 1):
            if not norm_enrollment(ws.cell(row, enroll_col).value):
                continue
            total += 1
            if to_number(ws.cell(row, col).value) is not None:
                valid += 1
        return valid, total

    if attended_col:
        valid, total = numeric_ratio(attended_col)
        if total == 0 or valid < max(1, int(total * 0.5)):
            attended_col = None

    if attended_col is None:
        for col in range(enroll_col + 1, ws.max_column + 1):
            valid, total = numeric_ratio(col)
            if total and valid >= max(1, int(total * 0.8)):
                attended_col = col
                break

    if attended_col is None:
        raise ValueError(
            f"Could not find the Classes Attended column in sheet '{ws.title}'."
        )

    return data_start, enroll_col, attended_col


def locate_register_columns(ws):
    try:
        header_row = find_header_row(ws, ["regno"])
    except ValueError:
        header_row = find_header_row(ws, ["rollno"])

    headers = [
        ws.cell(header_row, col).value
        for col in range(1, ws.max_column + 1)
    ]
    reg_col = find_column(headers, ["regno"])
    if not reg_col:
        reg_col = find_column(headers, ["enrollment", "rollno"])
    if not reg_col:
        raise ValueError("Register sheet must contain a REGNO or ROLLNO column.")

    total_lectures_col = find_total_lectures_column(headers)
    date_cols = [
        col for col, header in enumerate(headers, start=1)
        if is_date_header(header)
    ]
    if not date_cols:
        raise ValueError("No lecture-date columns were detected.")
    if not total_lectures_col:
        candidate = date_cols[-1] + 1
        if candidate <= ws.max_column:
            total_lectures_col = candidate

    return {
        "header_row": header_row,
        "headers": headers,
        "enrollment_col": reg_col,
        "date_cols": date_cols,
        "total_lectures_col": total_lectures_col,
        "total_attended_col": find_column(
            headers,
            [
                "total_lectures_attended",
                "total lectures attended",
                "lectures attended",
                "classes attended",
            ],
        ),
        "percentage_col": find_column(headers, ["percentage", "percent", "%"]),
    }


def analyze_workbooks(
    summary_file,
    register_file,
    summary_sheet_name,
    register_sheet_name,
):
    summary_wb = load_workbook(summary_file, data_only=True)
    register_wb = load_workbook(register_file, data_only=True)

    sws = summary_wb[summary_sheet_name]
    rws = register_wb[register_sheet_name]

    swapped_uploads = False
    if sheet_has_register_layout(sws) and not sheet_has_register_layout(rws):
        summary_wb, register_wb = register_wb, summary_wb
        sws, rws = rws, sws
        swapped_uploads = True

    summary_start, summary_enroll_col, summary_attended_col = locate_summary_columns(sws)
    summary_map = {}
    duplicate_summary = []
    for row in range(summary_start, sws.max_row + 1):
        enrollment = norm_enrollment(sws.cell(row, summary_enroll_col).value)
        if not enrollment:
            continue
        if enrollment in summary_map:
            duplicate_summary.append(enrollment)
        else:
            summary_map[enrollment] = to_number(
                sws.cell(row, summary_attended_col).value
            )

    register_details = locate_register_columns(rws)
    register_enrollments = []
    duplicate_register = []
    for row in range(register_details["header_row"] + 1, rws.max_row + 1):
        enrollment = norm_enrollment(
            rws.cell(row, register_details["enrollment_col"]).value
        )
        if not enrollment:
            continue
        if enrollment in register_enrollments:
            duplicate_register.append(enrollment)
        register_enrollments.append(enrollment)

    summary_set = set(summary_map)
    register_set = set(register_enrollments)
    date_cols = register_details["date_cols"]
    total_col = register_details["total_lectures_col"]
    total_values = []
    if total_col:
        for row in range(register_details["header_row"] + 1, rws.max_row + 1):
            value = to_number(rws.cell(row, total_col).value)
            if value is not None:
                total_values.append(int(round(value)))

    def column_label(col):
        return get_column_letter(col) if col else "Not detected"

    def column_header(col, header_row):
        if not col:
            return "Not detected"
        for row in range(max(header_row, 1), 0, -1):
            value = str(sws.cell(row, col).value or "").strip()
            if value:
                return value
        return "Not detected"

    if summary_start > 1:
        summary_attended_header = column_header(
            summary_attended_col, summary_start - 1
        )
        summary_enroll_header = column_header(
            summary_enroll_col, summary_start - 1
        )
    else:
        summary_attended_header = "Inferred from data"
        summary_enroll_header = "Inferred from data"

    return {
        "swapped_uploads": swapped_uploads,
        "summary": {
            "sheet": sws.title,
            "rows": sws.max_row,
            "students": len(summary_map),
            "enrollment_column": column_label(summary_enroll_col),
            "enrollment_header": summary_enroll_header,
            "attended_column": column_label(summary_attended_col),
            "attended_header": summary_attended_header,
            "duplicates": sorted(set(duplicate_summary)),
        },
        "register": {
            "sheet": rws.title,
            "rows": rws.max_row,
            "students": len(register_set),
            "enrollment_column": column_label(register_details["enrollment_col"]),
            "date_columns": len(date_cols),
            "first_date": str(register_details["headers"][date_cols[0] - 1] or ""),
            "last_date": str(register_details["headers"][date_cols[-1] - 1] or ""),
            "total_lectures_column": column_label(total_col),
            "total_attended_column": column_label(
                register_details["total_attended_col"]
            ),
            "percentage_column": column_label(
                register_details["percentage_col"]
            ),
            "total_lectures_values": sorted(set(total_values)),
            "duplicates": sorted(set(duplicate_register)),
        },
        "matched": len(summary_set & register_set),
        "summary_only": sorted(summary_set - register_set),
        "register_only": sorted(register_set - summary_set),
    }


def process_workbooks(
    summary_file,
    register_file,
    summary_sheet_name,
    register_sheet_name,
    seed=None,
):
    summary_wb = load_workbook(summary_file, data_only=True)
    register_wb = load_workbook(register_file)

    sws = summary_wb[summary_sheet_name]
    rws = register_wb[register_sheet_name]

    swapped_uploads = False
    if sheet_has_register_layout(sws) and not sheet_has_register_layout(rws):
        summary_wb, register_wb = register_wb, summary_wb
        sws, rws = rws, sws
        swapped_uploads = True

    summary_start, enroll_col, attended_col = locate_summary_columns(sws)

    summary_map = {}
    duplicate_summary = []
    for row in range(summary_start, sws.max_row + 1):
        enrollment = norm_enrollment(sws.cell(row, enroll_col).value)
        if not enrollment:
            continue
        attended = to_number(sws.cell(row, attended_col).value)
        if enrollment in summary_map:
            duplicate_summary.append(enrollment)
        else:
            summary_map[enrollment] = attended

    register_details = locate_register_columns(rws)
    reg_header = register_details["header_row"]
    reg_col = register_details["enrollment_col"]
    total_lectures_col = register_details["total_lectures_col"]
    date_cols = register_details["date_cols"]
    total_attended_col = register_details["total_attended_col"]
    percentage_col = register_details["percentage_col"]

    report = []
    errors = []
    matched = set()
    register_enrollments = []
    duplicate_register = []
    total_absences_marked = 0
    rng = random.Random(seed) if seed is not None else random.SystemRandom()

    for row in range(reg_header + 1, rws.max_row + 1):
        enrollment = norm_enrollment(rws.cell(row, reg_col).value)
        if not enrollment:
            continue
        if enrollment in register_enrollments:
            duplicate_register.append(enrollment)
        register_enrollments.append(enrollment)

        total_value = (
            to_number(rws.cell(row, total_lectures_col).value)
            if total_lectures_col
            else None
        )
        total_classes = (
            int(round(total_value))
            if total_value is not None and total_value > 0
            else len(date_cols)
        )

        if enrollment not in summary_map:
            report.append({
                "Enrollment No": enrollment,
                "Total Classes": total_classes,
                "Classes Attended": None,
                "Absent Classes": None,
                "Percentage": None,
                "Status": "NOT FOUND IN SUMMARY",
            })
            continue

        attended_value = summary_map[enrollment]
        if attended_value is None:
            errors.append(f"{enrollment}: classes attended is blank or not numeric in the summary.")
            report.append({
                "Enrollment No": enrollment,
                "Total Classes": total_classes,
                "Classes Attended": None,
                "Absent Classes": None,
                "Percentage": None,
                "Status": "INVALID ATTENDED VALUE",
            })
            continue

        attended = int(round(attended_value))
        absent = total_classes - attended
        if absent < 0:
            errors.append(
                f"{enrollment}: attended ({attended}) is greater than total classes ({total_classes})."
            )
            report.append({
                "Enrollment No": enrollment,
                "Total Classes": total_classes,
                "Classes Attended": attended,
                "Absent Classes": None,
                "Percentage": None,
                "Status": "INVALID TOTALS",
            })
            continue
        if absent > len(date_cols):
            errors.append(
                f"{enrollment}: absent ({absent}) exceeds available lecture dates ({len(date_cols)})."
            )
            report.append({
                "Enrollment No": enrollment,
                "Total Classes": total_classes,
                "Classes Attended": attended,
                "Absent Classes": None,
                "Percentage": None,
                "Status": "INVALID TOTALS",
            })
            continue

        for col in date_cols:
            cell = rws.cell(row, col)
            cell.value = None
            cell.fill = NO_FILL

        absent_cols = rng.sample(date_cols, absent)
        for col in absent_cols:
            cell = rws.cell(row, col)
            cell.value = "ABSENT"
            cell.fill = YELLOW_FILL

        percentage = round(attended / total_classes * 100, 2) if total_classes else 0

        if total_lectures_col:
            rws.cell(row, total_lectures_col).value = total_classes
        if total_attended_col:
            rws.cell(row, total_attended_col).value = attended
        if percentage_col:
            rws.cell(row, percentage_col).value = percentage

        matched.add(enrollment)
        total_absences_marked += absent
        report.append({
            "Enrollment No": enrollment,
            "Total Classes": total_classes,
            "Classes Attended": attended,
            "Absent Classes": absent,
            "Percentage": percentage,
            "Status": "MATCHED",
        })

    if errors:
        raise ValueError("Cannot mark attendance:\n" + "\n".join(errors[:20]))

    summary_only = sorted(set(summary_map) - set(register_enrollments))
    register_only = sorted(set(register_enrollments) - set(summary_map))

    output = io.BytesIO()
    register_wb.save(output)
    output.seek(0)

    stats = {
        "summary_students": len(summary_map),
        "register_students": len(set(register_enrollments)),
        "matched_students": len(matched),
        "summary_only": summary_only,
        "register_only": register_only,
        "duplicate_summary": sorted(set(duplicate_summary)),
        "duplicate_register": sorted(set(duplicate_register)),
        "date_columns": len(date_cols),
        "total_absences": total_absences_marked,
        "swapped_uploads": swapped_uploads,
        "report": pd.DataFrame(report),
    }
    return output.getvalue(), stats


def clean_output_name(name):
    name = (name or "").strip() or "Updated_Attendance_Register.xlsx"
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    if name.lower() == ".xlsx":
        name = "Updated_Attendance_Register.xlsx"
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def render_results(result, output_name=None):
    stats = result["stats"]
    report = stats["report"]
    matched = report[report["Status"] == "MATCHED"]

    st.divider()
    st.subheader("4. Results")

    metrics = st.columns(5)
    metrics[0].metric("Matched students", stats["matched_students"])
    metrics[1].metric("Absences marked", stats["total_absences"])
    metrics[2].metric(
        "Average attendance",
        f"{matched['Percentage'].mean():.2f}%" if not matched.empty else "N/A",
    )
    metrics[3].metric(
        "Below 75%",
        int((matched["Percentage"] < 75).sum()) if not matched.empty else 0,
    )
    metrics[4].metric("Lecture dates", stats["date_columns"])

    output_name = clean_output_name(output_name or result["output_name"])
    report_name = output_name[:-5] + "_processing_report.csv"
    downloads = st.columns(2)
    with downloads[0]:
        st.download_button(
            "Download updated Excel register",
            data=result["output_bytes"],
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            icon=":material/download:",
        )
    with downloads[1]:
        st.download_button(
            "Download processing report",
            data=report.to_csv(index=False).encode("utf-8-sig"),
            file_name=report_name,
            mime="text/csv",
            width="stretch",
            icon=":material/download:",
        )

    st.markdown("**Processing report**")
    st.dataframe(
        report,
        hide_index=True,
        alt="Student-level attendance processing report",
    )

    if stats["summary_only"]:
        st.warning(
            "Some students in the summary were not found in the register. "
            "They were not marked.",
            title="Unmatched summary students",
        )
    if stats["register_only"]:
        st.warning(
            "Some students in the register were not found in the summary. "
            "Their rows were left unchanged.",
            title="Unmatched register students",
        )
    if stats["duplicate_summary"] or stats["duplicate_register"]:
        st.error(
            "Duplicate enrollment numbers were detected. Verify the source workbooks.",
            title="Duplicate enrollment numbers",
        )

    with st.expander("Unmatched student details"):
        st.table(
            {
                "Summary only": ", ".join(stats["summary_only"]) or "None",
                "Register only": ", ".join(stats["register_only"]) or "None",
            },
            border="horizontal",
            width="content",
            alt="Unmatched students in the summary and register",
        )

    with st.expander("Audit details"):
        audit = {
            "Processed at": str(result["timestamp"]),
            "Summary workbook": str(result["summary_name"]),
            "Register workbook": str(result["register_name"]),
            "Summary worksheet": str(result["summary_sheet"]),
            "Register worksheet": str(result["register_sheet"]),
            "Random seed": (
                str(result["seed"])
                if result["seed"] is not None
                else "Not fixed"
            ),
            "Matched students": str(stats["matched_students"]),
            "Absences marked": str(stats["total_absences"]),
            "Lecture dates": str(stats["date_columns"]),
            "Uploaded in reverse order": (
                "Yes" if stats["swapped_uploads"] else "No"
            ),
        }
        st.table(
            audit,
            border="horizontal",
            width="content",
            alt="Attendance processing audit details",
        )

    st.info(
        "Verify a sample of marked students against the official attendance record "
        "before publishing."
    )


st.set_page_config(
    page_title="Attendance Register Automation",
    layout="wide",
    initial_sidebar_state="expanded",
)

with st.sidebar:
    st.subheader("Workflow")
    st.markdown(
        "1. Upload the summary workbook.\n"
        "2. Upload the detailed register.\n"
        "3. Review the detected columns.\n"
        "4. Process and download."
    )
    st.divider()
    st.caption(
        "Original files are not modified. Only matched students are updated. "
        "Unmatched students are reported and left unchanged."
    )

st.title("Attendance register automation")
st.caption(
    "Match a summary workbook with the detailed class register and produce a "
    "marked attendance register for faculty review."
)

with st.expander("How it works"):
    st.markdown(
        "1. Match students by enrollment number.\n"
        "2. Read classes attended from the summary and total classes from the register.\n"
        "3. Calculate absences as **total classes − classes attended**.\n"
        "4. Mark exactly that many random dates as **ABSENT** in yellow.\n"
        "5. Leave other dates blank and update totals and percentage."
    )

st.subheader("1. Upload workbooks")
st.caption(
    "Upload both `.xlsx` files. Reversed upload order is detected automatically."
)

left, right = st.columns(2)
with left:
    summary_upload = st.file_uploader(
        "Attendance summary workbook",
        type=["xlsx"],
        help="Should contain enrollment numbers and total classes attended.",
    )
    if summary_upload:
        st.caption(
            f"Selected: {summary_upload.name} "
            f"({summary_upload.size / 1024:.1f} KB)"
        )

with right:
    register_upload = st.file_uploader(
        "Detailed attendance register",
        type=["xlsx"],
        help="Should contain enrollment numbers, lecture-date columns, and total classes.",
    )
    if register_upload:
        st.caption(
            f"Selected: {register_upload.name} "
            f"({register_upload.size / 1024:.1f} KB)"
        )

if summary_upload and register_upload:
    summary_wb_preview = load_workbook(summary_upload, read_only=True, data_only=True)
    register_wb_preview = load_workbook(register_upload, read_only=True, data_only=True)
    sheet_columns = st.columns(2)
    with sheet_columns[0]:
        summary_sheet = st.selectbox(
            "Summary worksheet",
            summary_wb_preview.sheetnames,
            help="Select the worksheet that contains the attendance summary.",
        )
    with sheet_columns[1]:
        register_sheet = st.selectbox(
            "Register worksheet",
            register_wb_preview.sheetnames,
            help="Select the worksheet that contains the detailed class register.",
        )
    summary_wb_preview.close()
    register_wb_preview.close()

    summary_upload.seek(0)
    register_upload.seek(0)

    analysis = None
    try:
        analysis = analyze_workbooks(
            summary_upload,
            register_upload,
            summary_sheet,
            register_sheet,
        )
        summary_upload.seek(0)
        register_upload.seek(0)
    except Exception as exc:
        st.error(f"Could not inspect the uploaded workbooks: {exc}")

    if analysis:
        st.subheader("2. Review detected structure")
        if analysis["swapped_uploads"]:
            st.info(
                "The workbooks were uploaded in reverse order. "
                "Their roles were detected automatically."
            )

        structure = st.columns(2)
        with structure[0]:
            st.markdown("**Attendance summary**")
            st.table(
                {
                    "Worksheet": str(analysis["summary"]["sheet"]),
                    "Students": str(analysis["summary"]["students"]),
                    "Enrollment column": (
                        f"{analysis['summary']['enrollment_column']} "
                        f"({analysis['summary']['enrollment_header']})"
                    ),
                    "Classes attended": (
                        f"{analysis['summary']['attended_column']} "
                        f"({analysis['summary']['attended_header']})"
                    ),
                },
                border="horizontal",
                width="stretch",
                alt="Detected attendance summary structure",
            )

        with structure[1]:
            register_info = analysis["register"]
            totals_text = (
                ", ".join(str(value) for value in register_info["total_lectures_values"])
                if register_info["total_lectures_values"]
                else "Not detected"
            )
            st.markdown("**Detailed attendance register**")
            st.table(
                {
                    "Worksheet": str(register_info["sheet"]),
                    "Students": str(register_info["students"]),
                    "Enrollment column": str(register_info["enrollment_column"]),
                    "Lecture dates": (
                        f"{register_info['date_columns']} "
                        f"({register_info['first_date']} to {register_info['last_date']})"
                    ),
                    "Total classes": (
                        f"{register_info['total_lectures_column']} "
                        f"(values: {totals_text})"
                    ),
                    "Total attended": str(register_info["total_attended_column"]),
                    "Percentage": str(register_info["percentage_column"]),
                },
                border="horizontal",
                width="stretch",
                alt="Detected detailed register structure",
            )

        blocking_issues = []
        if analysis["summary"]["students"] == 0:
            blocking_issues.append("No enrollment numbers were found in the summary.")
        if analysis["register"]["students"] == 0:
            blocking_issues.append("No enrollment numbers were found in the register.")
        if analysis["register"]["date_columns"] == 0:
            blocking_issues.append("No lecture-date columns were detected.")
        if analysis["matched"] == 0:
            blocking_issues.append("No students matched between the two workbooks.")
        if analysis["summary"]["duplicates"]:
            blocking_issues.append(
                "Duplicate enrollment numbers were found in the summary."
            )
        if analysis["register"]["duplicates"]:
            blocking_issues.append(
                "Duplicate enrollment numbers were found in the register."
            )

        if blocking_issues:
            st.error(
                "\n".join(f"- {issue}" for issue in blocking_issues),
                title="Resolve these issues before processing",
            )
        else:
            st.success(
                f"{analysis['matched']} students matched across "
                f"{analysis['register']['date_columns']} lecture dates."
            )

        if analysis["summary_only"]:
            st.warning(
                f"{len(analysis['summary_only'])} summary student(s) were not found "
                "in the register and will not be marked.",
                title="Unmatched summary students",
            )
        if analysis["register_only"]:
            st.warning(
                f"{len(analysis['register_only'])} register student(s) were not found "
                "in the summary and will remain unchanged.",
                title="Unmatched register students",
            )

        st.subheader("3. Processing options")
        options = st.columns([1, 1])
        with options[0]:
            seed_text = st.text_input(
                "Random seed (optional)",
                value="",
                help=(
                    "Use a number to reproduce the same random absence dates. "
                    "Leave blank for a new random selection each run."
                ),
            )
        with options[1]:
            output_name_input = st.text_input(
                "Output file name",
                value="Updated_Attendance_Register.xlsx",
                help="The downloaded Excel file will use this name.",
            )

        confirmed = st.checkbox(
            "I have reviewed the detected columns and confirm that absence dates "
            "should be assigned randomly."
        )

        signature = (
            summary_upload.name,
            summary_upload.size,
            register_upload.name,
            register_upload.size,
            summary_sheet,
            register_sheet,
            seed_text.strip(),
        )

        process_ready = not blocking_issues and analysis["matched"] > 0
        if st.button(
            "Process attendance register",
            type="primary",
            width="stretch",
            disabled=not (process_ready and confirmed),
        ):
            try:
                st.session_state.pop("attendance_result", None)
                summary_upload.seek(0)
                register_upload.seek(0)
                seed = int(seed_text) if seed_text.strip() else None
                with st.spinner("Processing attendance records..."):
                    output_bytes, stats = process_workbooks(
                        summary_upload,
                        register_upload,
                        summary_sheet,
                        register_sheet,
                        seed=seed,
                    )
                st.session_state["attendance_result"] = {
                    "signature": signature,
                    "output_bytes": output_bytes,
                    "stats": stats,
                    "output_name": clean_output_name(output_name_input),
                    "summary_name": summary_upload.name,
                    "register_name": register_upload.name,
                    "summary_sheet": summary_sheet,
                    "register_sheet": register_sheet,
                    "seed": seed,
                    "timestamp": datetime.now().strftime("%d %b %Y, %H:%M"),
                }
                st.success(
                    "Attendance register processed successfully. "
                    "The original workbooks were not modified."
                )
            except Exception as exc:
                st.error(f"Could not process the workbooks: {exc}")

        if not process_ready:
            st.caption(
                "Processing is disabled until the blocking issues above are resolved."
            )
        elif not confirmed:
            st.caption(
                "Confirm the review checkbox to enable processing."
            )

        current_signature = signature
        stored_result = st.session_state.get("attendance_result")
        if stored_result and stored_result["signature"] == current_signature:
            render_results(stored_result, output_name_input)
        elif stored_result:
            st.info(
                "The uploaded files or processing options changed. "
                "Run the processing step again to refresh the results."
            )
else:
    st.info("Upload both Excel workbooks to begin the review.")

st.caption(
    "Verify the generated register against the official attendance record "
    "before publication."
)
