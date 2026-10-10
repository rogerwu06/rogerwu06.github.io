# -*- coding: utf-8 -*-
"""
Abaqus Pressure Sweep AutoRunner

Normal Python mode:
    - asks the user to manually select the Abaqus CAE launcher/shortcut
    - asks the user to select the Abaqus .cae file
    - creates a new timestamped run folder beside the selected CAE
    - copies the CAE into the new run folder
    - creates a Jobs subfolder for collected Abaqus job/result files
    - visibly opens the copied CAE in Abaqus/CAE and runs this same script inside Abaqus
    - when Abaqus finishes, creates the final Excel workbook and single-series graph

Abaqus worker mode:
    - uses the copied CAE that Abaqus already opened
    - finds the existing pressure jobs already stored in the CAE
    - submits those existing jobs sequentially without rebuilding them
    - extracts max Mises stress and max PEEQ from the last available ODB frame
    - stops immediately when a job does not complete
    - on failure, reads the last STEP TIME/LPF value from the .sta monitor/status file
    - calculates adjusted failure pressure = attempted pressure * STEP TIME/LPF
    - writes a CSV after every pressure so completed results are retained

The Abaqus execution/submission logic is intentionally kept the same as the working
v4 version. Only output collection and post-processing are changed here. The final
Excel/PNG post-processing is done by normal Python because normal Python is much
more likely to have matplotlib, SciPy, and openpyxl available.
"""

from __future__ import print_function

import os
import sys
import csv
import time
import shutil
import subprocess
import glob


##
# main part
#

START_PRESSURE = 6
END_PRESSURE = 120
PRESSURE_STEP = 1

STEP_NAME = 'Step-1'
FRAME_NUMBER = -1                 # -1 means use the last available frame

# abaqus stuff
# files here

DESIGN_PRESSURE = 9.5

CSV_NAME = 'PEEQ_Stress_Pressure_Summary.csv'
XLSX_NAME = 'PEEQ_Stress_Pressure_Summary.xlsx'
GRAPH_NAME = 'failure_pressure_prediction.png'
RUN_INFO_NAME = 'Run_Information.txt'

# abaqus run
# model stuff
# file stuff
LAUNCH_DIAGNOSTICS_NAME = 'Abaqus_Launch_Diagnostics.txt'
WORKER_DONE_NAME = 'Abaqus_Worker_Complete.flag'
WORKER_ERROR_NAME = 'Abaqus_Worker_Error.txt'

# abaqus part
ENV_WORKER_MODE = 'ABAQUS_PRESSURE_SWEEP_WORKER'
ENV_CAE_PATH = 'ABAQUS_PRESSURE_SWEEP_CAE'
ENV_RUN_ROOT = 'ABAQUS_PRESSURE_SWEEP_ROOT'
ENV_JOBS_DIR = 'ABAQUS_PRESSURE_SWEEP_JOBS'
ENV_SOURCE_DIR = 'ABAQUS_PRESSURE_SWEEP_SOURCE_DIR'


#
# leave this like this
####

def safe_float(value):
    try:
        return float(str(value).replace('D', 'E').replace('d', 'e'))
    except:
        return None


def pressure_text(value):
    value = float(value)
    if abs(value - round(value)) < 1.0e-10:
        return str(int(round(value)))
    return ('%.8f' % value).rstrip('0').rstrip('.')


#
# main part
#
def timestamp_text():
    return time.strftime('%Y%m%d_%H%M%S')


def unique_folder(path):
    if not os.path.exists(path):
        return path

    index = 2
    while True:
        candidate = path + '_%d' % index
        if not os.path.exists(candidate):
            return candidate
        index += 1


def write_run_information(run_root, source_cae, copied_cae, jobs_dir):
    info_path = os.path.join(run_root, RUN_INFO_NAME)
    handle = open(info_path, 'w')
    try:
        handle.write('ABAQUS AUTOMATIC PRESSURE SWEEP RUN\n')
        handle.write('===================================\n\n')
        handle.write('Source CAE: %s\n' % source_cae)
        handle.write('Copied CAE used for run: %s\n' % copied_cae)
        handle.write('Jobs folder: %s\n' % jobs_dir)
        handle.write('Start pressure: %s MPa\n' % START_PRESSURE)
        handle.write('End pressure: %s MPa\n' % END_PRESSURE)
        handle.write('Pressure step: %s MPa\n' % PRESSURE_STEP)
        handle.write('Result step: %s\n' % STEP_NAME)
        handle.write('Result frame: %s\n' % FRAME_NUMBER)
        handle.write('\nFailure rule:\n')
        handle.write('On first non-completed Abaqus job, stop the sweep.\n')
        handle.write('Adjusted failure pressure = attempted pressure x last STEP TIME/LPF.\n')
    finally:
        handle.close()


#
# path stuff
#

####
# this part
####
def choose_abaqus_file():
    """Ask the user to manually select the Abaqus launcher used for this run."""
    try:
        try:
            import tkinter as tk
            from tkinter import filedialog
        except ImportError:
            import Tkinter as tk
            import tkFileDialog as filedialog

        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes('-topmost', True)
        except:
            pass

        initial_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
        if not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser('~')

        selected = filedialog.askopenfilename(
            title='FIRST: Select Abaqus CAE launcher or Abaqus CAE shortcut',
            initialdir=initial_dir,
            filetypes=[
                ('Abaqus launcher / shortcut', '*.lnk *.bat *.cmd *.exe'),
                ('Windows shortcut', '*.lnk'),
                ('Batch launcher', '*.bat *.cmd'),
                ('Executable', '*.exe'),
                ('All files', '*.*')
            ]
        )

        try:
            root.destroy()
        except:
            pass

        if selected:
            return os.path.abspath(selected)

    except Exception as exc:
        print('Could not open the Abaqus picker: %s' % exc)

    try:
        user_value = input('Enter the full path to the Abaqus launcher/shortcut: ').strip()
    except NameError:
        user_value = raw_input('Enter the full path to the Abaqus launcher/shortcut: ').strip()

    user_value = user_value.strip('"')
    if user_value:
        return os.path.abspath(user_value)

    return None


def choose_cae_file():
    # file stuff
    if len(sys.argv) > 1:
        candidate = os.path.abspath(sys.argv[1])
        if os.path.isfile(candidate) and candidate.lower().endswith('.cae'):
            return candidate

    try:
        try:
            import tkinter as tk
            from tkinter import filedialog
        except ImportError:
            import Tkinter as tk
            import tkFileDialog as filedialog

        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes('-topmost', True)
        except:
            pass

        selected = filedialog.askopenfilename(
            title='Select the Abaqus CAE file to run',
            filetypes=[('Abaqus CAE files', '*.cae'), ('All files', '*.*')]
        )

        try:
            root.destroy()
        except:
            pass

        if selected:
            return os.path.abspath(selected)

    except Exception as exc:
        print('Could not open the file picker: %s' % exc)

    # leave this like this
    try:
        user_value = input('Enter the full path to the Abaqus .cae file: ').strip()
    except NameError:
        user_value = raw_input('Enter the full path to the Abaqus .cae file: ').strip()

    user_value = user_value.strip('"')
    if user_value:
        return os.path.abspath(user_value)

    return None


#
# save files
#

def read_results_csv(csv_path):
    results = []

    if not os.path.exists(csv_path):
        return results

    handle = open(csv_path, 'r')
    try:
        reader = csv.DictReader(handle)
        for row in reader:
            results.append({
                'Attempted Pressure (MPa)': safe_float(row.get('Attempted Pressure (MPa)')),
                'Step Time/LPF': safe_float(row.get('Step Time/LPF')),
                'Effective Pressure (MPa)': safe_float(row.get('Effective Pressure (MPa)')),
                'Job Name': row.get('Job Name', ''),
                'Status': row.get('Status', ''),
                'Max Mises Stress': safe_float(row.get('Max Mises Stress')),
                'Max PEEQ': safe_float(row.get('Max PEEQ')),
                'Failure Stop': row.get('Failure Stop', '')
            })
    finally:
        handle.close()

    return results


##
# needed here
#

#
# main part
#
def clean_xy(x_values, y_values):
    pairs = []
    for x_value, y_value in zip(x_values, y_values):
        if x_value is None or y_value is None:
            continue
        pairs.append((float(x_value), float(y_value)))

    pairs.sort(key=lambda item: item[0])

    # input stuff
    # inputs here
    cleaned = []
    for x_value, y_value in pairs:
        if cleaned and abs(cleaned[-1][0] - x_value) < 1.0e-12:
            cleaned[-1] = (x_value, y_value)
        else:
            cleaned.append((x_value, y_value))

    return [item[0] for item in cleaned], [item[1] for item in cleaned]


def add_tangent_arrows(ax, x_data, y_data, np_module):
    n = len(x_data)
    if n < 3:
        return

    base_length = 1.6

    def slope(index):
        if index == 0:
            return (y_data[1] - y_data[0]) / (x_data[1] - x_data[0])
        elif index == n - 1:
            return (y_data[-1] - y_data[-2]) / (x_data[-1] - x_data[-2])
        else:
            return ((y_data[index + 1] - y_data[index - 1]) /
                    (x_data[index + 1] - x_data[index - 1]))

    indices = [1, n // 2, n - 2]
    used = []
    for index in indices:
        if index in used or index < 0 or index >= n:
            continue
        used.append(index)

        arrow_length = base_length
        if index == n - 2:
            arrow_length *= 1.5

        m_value = slope(index)
        dx = 1.0
        dy = m_value
        magnitude = np_module.sqrt(dx ** 2 + dy ** 2)

        if magnitude == 0:
            continue

        dx = dx / magnitude * arrow_length
        dy = dy / magnitude * arrow_length

        ax.annotate(
            '',
            xy=(x_data[index] + dx, y_data[index] + dy),
            xytext=(x_data[index] - dx, y_data[index] - dy),
            arrowprops=dict(
                arrowstyle='->',
                linewidth=1.2,
                color='black',
                alpha=0.5
            ),
            zorder=10
        )


def create_graph(results, graph_path, model_label):
    """
    Same general appearance as graphcreatorforpressuretests.py, but ONLY the
    actual pressure/PEEQ results from this run are plotted. There is no
    reinforced/unreinforced comparison dataset.
    """
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import numpy as np
        from matplotlib.lines import Line2D
    except ImportError as exc:
        print('Graph was not created because a plotting package is missing: %s' % exc)
        return None

    try:
        from scipy.interpolate import PchipInterpolator
        have_pchip = True
    except ImportError:
        PchipInterpolator = None
        have_pchip = False

    x_data = []
    y_data = []

    for row in results:
        x_value = row.get('Effective Pressure (MPa)')
        y_value = row.get('Max PEEQ')
        if x_value is not None and y_value is not None:
            x_data.append(x_value)
            y_data.append(y_value)

    x_data, y_data = clean_xy(x_data, y_data)

    if not x_data:
        print('No PEEQ data was available, so the graph was not created.')
        return None

    plt.rcParams['font.family'] = 'Calibri'
    fig, ax = plt.subplots(figsize=(15, 6))
    blue = '#4472C4'

    # screen layout
    if have_pchip and len(x_data) >= 2:
        dense_x = np.linspace(min(x_data), max(x_data), 400)
        curve = PchipInterpolator(x_data, y_data)
        model_line, = ax.plot(
            dense_x,
            curve(dense_x),
            color=blue,
            linewidth=2.5,
            label=model_label,
            zorder=3
        )
    else:
        model_line, = ax.plot(
            x_data,
            y_data,
            color=blue,
            linewidth=2.5,
            label=model_label,
            zorder=3
        )

    ax.plot(x_data, y_data, 'o', color=blue, markersize=5, zorder=4)
    add_tangent_arrows(ax, x_data, y_data, np)

    ax.axvline(
        DESIGN_PRESSURE,
        color='#7F7F7F',
        linestyle='--',
        linewidth=1.5
    )

    # stuff below
    # results
    y_max = max(25.0, max(y_data) * 1.15 if y_data else 25.0)
    x_max = max(20.0, max(x_data + [DESIGN_PRESSURE]) * 1.10)

    ax.text(
        DESIGN_PRESSURE - 0.5,
        y_max * 0.88,
        'Design Pressure',
        fontsize=10,
        bbox=dict(
            boxstyle='round',
            facecolor='white',
            edgecolor='none'
        )
    )

    arrow_handle = Line2D(
        [0],
        [0],
        color='black',
        linewidth=1.2,
        marker=r'$\rightarrow$',
        markersize=12,
        label='Tangent Direction'
    )

    ax.legend(
        handles=[model_line, arrow_handle],
        loc='upper left',
        frameon=False
    )

    ax.set_xlabel('Applied Pressure (MPa)', fontsize=12)
    ax.set_ylabel('PEEQ (%)', fontsize=12)
    ax.set_title('Failure Pressure Prediction', fontsize=14)
    ax.grid(True, color='#D9D9D9')
    ax.set_axisbelow(True)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_xlim(0, x_max)
    ax.set_ylim(0, y_max)

    plt.tight_layout()
    plt.savefig(graph_path, dpi=300, bbox_inches='tight')
    plt.close(fig)

    print('Graph saved: %s' % graph_path)
    return graph_path


# #
# path stuff
##

####
# this part
####
def create_excel(results, xlsx_path, graph_path, model_label):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
        from openpyxl.drawing.image import Image as ExcelImage
    except ImportError as exc:
        print('Excel workbook was not created because openpyxl is missing: %s' % exc)
        return None

    wb = Workbook()

    ws = wb.active
    ws.title = 'Results'

    headers = [
        'Attempted Pressure (MPa)',
        'Step Time/LPF',
        'Effective Pressure (MPa)',
        'Job Name',
        'Status',
        'Max Mises Stress',
        'Max PEEQ',
        'Failure Stop'
    ]

    for col_index, heading in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_index, value=heading)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for row_index, row in enumerate(results, 2):
        values = [
            row.get('Attempted Pressure (MPa)'),
            row.get('Step Time/LPF'),
            row.get('Effective Pressure (MPa)'),
            row.get('Job Name'),
            row.get('Status'),
            row.get('Max Mises Stress'),
            row.get('Max PEEQ'),
            row.get('Failure Stop')
        ]

        for col_index, value in enumerate(values, 1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            cell.alignment = Alignment(horizontal='center')

    widths = [24, 16, 25, 24, 16, 20, 16, 16]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width

    summary = wb.create_sheet('Failure Summary')
    summary['A1'] = 'Pressure Sweep Summary'
    summary['A1'].font = Font(bold=True, size=14)
    summary['A2'] = 'Model'
    summary['B2'] = model_label

    failure_rows = [row for row in results if str(row.get('Failure Stop', '')).upper() == 'YES']

    summary['A3'] = 'Last requested pressure'
    summary['B3'] = results[-1].get('Attempted Pressure (MPa)') if results else None

    if failure_rows:
        failed = failure_rows[-1]
        summary['A4'] = 'First failed attempted pressure (MPa)'
        summary['B4'] = failed.get('Attempted Pressure (MPa)')
        summary['A5'] = 'Last STEP TIME/LPF'
        summary['B5'] = failed.get('Step Time/LPF')
        summary['A6'] = 'Adjusted failure pressure (MPa)'
        summary['B6'] = failed.get('Effective Pressure (MPa)')
        summary['A7'] = 'PEEQ at last available frame'
        summary['B7'] = failed.get('Max PEEQ')
        summary['A8'] = 'Mises stress at last available frame'
        summary['B8'] = failed.get('Max Mises Stress')
        summary['A10'] = 'Calculation'
        summary['B10'] = 'Attempted pressure x STEP TIME/LPF'
    else:
        summary['A4'] = 'Failure detected'
        summary['B4'] = 'No'

    summary.column_dimensions['A'].width = 38
    summary.column_dimensions['B'].width = 48

    # main part
    data_ws = wb.create_sheet('Graph Data')
    graph_headers = [
        model_label + ' Pressure (MPa)',
        model_label + ' PEEQ'
    ]

    for col_index, heading in enumerate(graph_headers, 1):
        cell = data_ws.cell(row=1, column=col_index, value=heading)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    graph_rows = []
    for row in results:
        if row.get('Effective Pressure (MPa)') is not None and row.get('Max PEEQ') is not None:
            graph_rows.append((row.get('Effective Pressure (MPa)'), row.get('Max PEEQ')))

    for index, pair in enumerate(graph_rows, 2):
        data_ws.cell(row=index, column=1, value=pair[0])
        data_ws.cell(row=index, column=2, value=pair[1])

    data_ws.column_dimensions['A'].width = max(24, min(55, len(model_label) + 18))
    data_ws.column_dimensions['B'].width = max(24, min(55, len(model_label) + 10))

    graph_ws = wb.create_sheet('Failure Pressure Graph')
    graph_ws['A1'] = 'Failure Pressure Prediction - ' + model_label
    graph_ws['A1'].font = Font(bold=True, size=14)

    if graph_path and os.path.exists(graph_path):
        try:
            image = ExcelImage(graph_path)
            image.width = 1200
            image.height = 480
            graph_ws.add_image(image, 'A3')
        except Exception as exc:
            graph_ws['A3'] = 'Graph image could not be embedded: %s' % exc

    wb.save(xlsx_path)
    print('Excel workbook saved: %s' % xlsx_path)
    return xlsx_path


#
# abaqus run
#

def _decode_console_output(value):
    if value is None:
        return ''
    if isinstance(value, bytes):
        for encoding in ('utf-8', 'mbcs', 'latin-1'):
            try:
                return value.decode(encoding)
            except:
                pass
        return value.decode('utf-8', 'replace')
    return str(value)


def _powershell_single_quote(value):
    return str(value).replace("'", "''")


#
# main part
#
def resolve_windows_shortcut(shortcut_path):
    """Return shortcut target/arguments/working directory using WScript.Shell."""
    if not shortcut_path or not os.path.isfile(shortcut_path):
        return None

    ps_path = _powershell_single_quote(shortcut_path)
    ps_code = (
        "$w = New-Object -ComObject WScript.Shell; "
        "$s = $w.CreateShortcut('%s'); "
        "Write-Output ('TARGET=' + $s.TargetPath); "
        "Write-Output ('ARGS=' + $s.Arguments); "
        "Write-Output ('WORKDIR=' + $s.WorkingDirectory)"
    ) % ps_path

    try:
        process = subprocess.Popen(
            [
                'powershell.exe',
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-Command', ps_code
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        stdout, stderr = process.communicate()
        stdout = _decode_console_output(stdout)
        stderr = _decode_console_output(stderr)
    except Exception as exc:
        return {
            'target': '',
            'arguments': '',
            'working_directory': '',
            'error': 'PowerShell shortcut resolution failed: %s' % exc
        }

    data = {
        'target': '',
        'arguments': '',
        'working_directory': '',
        'error': ''
    }

    for raw_line in stdout.splitlines():
        line = raw_line.strip()
        if line.startswith('TARGET='):
            data['target'] = line[len('TARGET='):].strip()
        elif line.startswith('ARGS='):
            data['arguments'] = line[len('ARGS='):].strip()
        elif line.startswith('WORKDIR='):
            data['working_directory'] = line[len('WORKDIR='):].strip()

    if process.returncode != 0:
        data['error'] = 'PowerShell returned %s: %s' % (process.returncode, stderr.strip())
    elif not data['target']:
        data['error'] = 'The shortcut did not report a target path. %s' % stderr.strip()

    return data


def _extract_launcher_from_shortcut(shortcut_data):
    """
    Abaqus desktop shortcuts sometimes point directly to abq2023.bat, but some
    point to cmd.exe and put the real Abaqus .bat file inside the shortcut args.
    This tries both layouts and returns the real launcher path when possible.
    """
    if not shortcut_data:
        return None

    target = shortcut_data.get('target', '') or ''
    arguments = shortcut_data.get('arguments', '') or ''

    target_lower = target.lower()
    target_base = os.path.basename(target_lower)

    if target and target_base not in ('cmd.exe', 'powershell.exe', 'pwsh.exe'):
        if target_lower.endswith(('.bat', '.cmd', '.exe')):
            return target

    # this part
    # abaqus part
    import re
    quoted = re.findall(r'[\"\']([^\"\']+\.(?:bat|cmd|exe))[\"\']', arguments, flags=re.I)
    for candidate in quoted:
        if os.path.isfile(candidate):
            return candidate

    # path stuff
    # model stuff
    tokens = re.findall(r'([A-Za-z]:\\[^\r\n]+?\.(?:bat|cmd|exe))(?=\s|$)', arguments, flags=re.I)
    for candidate in tokens:
        candidate = candidate.strip(' \"\'')
        if os.path.isfile(candidate) and ('abq' in os.path.basename(candidate).lower() or
                                          'abaqus' in os.path.basename(candidate).lower()):
            return candidate

    return None


def resolve_selected_abaqus_launcher(selected_path):
    """Resolve the exact Abaqus launcher selected by the user."""
    info = {
        'selected_path': selected_path or '',
        'shortcut': '',
        'shortcut_target': '',
        'shortcut_arguments': '',
        'shortcut_working_directory': '',
        'launcher': '',
        'resolution_error': ''
    }

    if not selected_path:
        info['resolution_error'] = 'No Abaqus launcher was selected.'
        return info

    selected_path = os.path.abspath(selected_path)
    info['selected_path'] = selected_path

    if not os.path.isfile(selected_path):
        info['resolution_error'] = 'Selected Abaqus file does not exist: %s' % selected_path
        return info

    lower = selected_path.lower()

    if lower.endswith('.lnk'):
        info['shortcut'] = selected_path
        shortcut_data = resolve_windows_shortcut(selected_path)
        if not shortcut_data:
            info['resolution_error'] = 'Could not resolve the selected Windows shortcut.'
            return info

        info['shortcut_target'] = shortcut_data.get('target', '')
        info['shortcut_arguments'] = shortcut_data.get('arguments', '')
        info['shortcut_working_directory'] = shortcut_data.get('working_directory', '')
        info['resolution_error'] = shortcut_data.get('error', '')

        launcher = _extract_launcher_from_shortcut(shortcut_data)
        if launcher:
            info['launcher'] = launcher
            return info

        if not info['resolution_error']:
            info['resolution_error'] = (
                'The shortcut was resolved, but the real Abaqus .bat/.cmd/.exe '
                'could not be extracted. Select the Abaqus command launcher '
                '(for example abq2023.bat) directly instead.'
            )
        return info

    if lower.endswith(('.bat', '.cmd', '.exe')):
        info['launcher'] = selected_path
        return info

    info['resolution_error'] = (
        'Unsupported Abaqus selection. Select Abaqus CAE.lnk, abq2023.bat, '
        'another Abaqus .bat/.cmd launcher, or an Abaqus executable.'
    )
    return info

####
# this part
####
def build_abaqus_command(launcher, database_path, worker_script, visible=True):
    """Open the selected CAE first, then run the automation as an Abaqus/CAE script."""
    launcher_text = str(launcher).strip().strip('"')
    database_text = os.path.abspath(database_path)
    worker_text = os.path.abspath(worker_script)

    # abaqus stuff
    # abaqus part
    if visible:
        abaqus_args = 'cae database="{}" script="{}"'.format(database_text, worker_text)
    else:
        # this part
        abaqus_args = 'cae noGUI="{}" -- "{}"'.format(worker_text, database_text)

    if launcher_text.lower().endswith(('.bat', '.cmd')):
        return 'call "{}" {}'.format(launcher_text, abaqus_args)

    return '"{}" {}'.format(launcher_text, abaqus_args)

def write_launch_diagnostics(run_root, launch_info, command, return_code=None, exception_text=None):
    path = os.path.join(run_root, LAUNCH_DIAGNOSTICS_NAME)
    try:
        handle = open(path, 'w')
        try:
            handle.write('ABAQUS LAUNCH DIAGNOSTICS\n')
            handle.write('=========================\n\n')
            handle.write('Manually selected Abaqus file: %s\n' % launch_info.get('selected_path', ''))
            handle.write('Selected shortcut: %s\n' % launch_info.get('shortcut', ''))
            handle.write('Shortcut target: %s\n' % launch_info.get('shortcut_target', ''))
            handle.write('Shortcut arguments: %s\n' % launch_info.get('shortcut_arguments', ''))
            handle.write('Shortcut working directory: %s\n' % launch_info.get('shortcut_working_directory', ''))
            handle.write('Resolved launcher: %s\n' % launch_info.get('launcher', ''))
            handle.write('Launch mode: visible Abaqus/CAE script mode\n')
            handle.write('Resolution warning/error: %s\n' % launch_info.get('resolution_error', ''))
            handle.write('\nCommand executed:\n%s\n' % command)
            handle.write('\nReturn code: %s\n' % str(return_code))
            if exception_text:
                handle.write('Launch exception: %s\n' % exception_text)
        finally:
            handle.close()
    except:
        pass
    return path


#
# abaqus stuff
#

def run_normal_python_launcher():
    print('')
    print('=' * 78)
    print('ABAQUS PRESSURE SWEEP - MANUAL ABAQUS SELECTION')
    print('=' * 78)
    print('1) Select the Abaqus CAE launcher/shortcut.')
    print('2) Select the CAE file to run.')
    print('3) Abaqus/CAE will open visibly and the automation will run inside it.')
    print('=' * 78)

    selected_abaqus = choose_abaqus_file()
    if not selected_abaqus:
        print('No Abaqus launcher was selected. Nothing was run.')
        return 1

    launch_info = resolve_selected_abaqus_launcher(selected_abaqus)
    launcher = launch_info.get('launcher')

    if not launcher:
        print('')
        print('Could not determine the real Abaqus launcher from your selection.')
        print(launch_info.get('resolution_error', 'Unknown launcher error.'))
        print('')
        print('If you selected Abaqus CAE.lnk, try selecting the real Abaqus')
        print('command file instead, usually named something like abq2023.bat.')
        return 1

    source_cae = choose_cae_file()

    if not source_cae:
        print('No CAE file was selected. Nothing was run.')
        return 1

    if not os.path.isfile(source_cae):
        print('CAE file does not exist: %s' % source_cae)
        return 1

    if not source_cae.lower().endswith('.cae'):
        print('Selected file is not a .cae file: %s' % source_cae)
        return 1

    source_directory = os.path.dirname(source_cae)
    base_name = os.path.splitext(os.path.basename(source_cae))[0]

    run_folder_name = base_name + '_PressureSweep_' + timestamp_text()
    run_root = unique_folder(os.path.join(source_directory, run_folder_name))
    jobs_dir = os.path.join(run_root, 'Jobs')

    os.makedirs(run_root)
    os.makedirs(jobs_dir)

    copied_cae = os.path.join(run_root, os.path.basename(source_cae))
    shutil.copy2(source_cae, copied_cae)

    try:
        original_script = os.path.abspath(__file__)
        worker_script = os.path.join(run_root, 'Automation_Script_Used.py')
        shutil.copy2(original_script, worker_script)
    except Exception:
        worker_script = os.path.abspath(__file__)

    write_run_information(run_root, source_cae, copied_cae, jobs_dir)

    done_flag = os.path.join(run_root, WORKER_DONE_NAME)
    error_file = os.path.join(run_root, WORKER_ERROR_NAME)
    for stale_path in (done_flag, error_file):
        try:
            if os.path.exists(stale_path):
                os.remove(stale_path)
        except:
            pass

    print('')
    print('=' * 78)
    print('ABAQUS PRESSURE SWEEP')
    print('=' * 78)
    print('Selected Abaqus file: %s' % selected_abaqus)
    if launch_info.get('shortcut_target'):
        print('Shortcut target: %s' % launch_info.get('shortcut_target'))
    if launch_info.get('shortcut_arguments'):
        print('Shortcut arguments: %s' % launch_info.get('shortcut_arguments'))
    print('Resolved Abaqus launcher: %s' % launcher)
    print('Source CAE: %s' % source_cae)
    print('Run folder: %s' % run_root)
    print('Jobs folder: %s' % jobs_dir)
    print('Copied CAE: %s' % copied_cae)
    print('=' * 78)

    env = os.environ.copy()
    env[ENV_WORKER_MODE] = '1'
    env[ENV_CAE_PATH] = copied_cae
    env[ENV_RUN_ROOT] = run_root
    env[ENV_JOBS_DIR] = jobs_dir
    # abaqus run
    # input values
    # abaqus stuff
    # files here
    env[ENV_SOURCE_DIR] = source_directory

    command = build_abaqus_command(launcher, copied_cae, worker_script, visible=True)

    print('')
    print('Opening visible Abaqus/CAE and running the worker script...')
    print(command)
    print('')

    diagnostics_path = write_launch_diagnostics(run_root, launch_info, command)

    try:
        process = subprocess.Popen(command, cwd=run_root, env=env, shell=True)
    except Exception as exc:
        write_launch_diagnostics(run_root, launch_info, command, exception_text=str(exc))
        print('Could not launch Abaqus: %s' % exc)
        print('Launch diagnostics: %s' % diagnostics_path)
        return 1

    # abaqus run
    # stuff below
    # gui stuff
    last_message_time = 0.0
    worker_finished = False
    process_return_code = None

    while True:
        if os.path.exists(done_flag):
            worker_finished = True
            break

        if os.path.exists(error_file):
            print('')
            print('The Abaqus worker reported an error:')
            try:
                handle = open(error_file, 'r')
                try:
                    print(handle.read())
                finally:
                    handle.close()
            except:
                pass
            break

        process_return_code = process.poll()
        if process_return_code is not None:
            # leave this like this
            time.sleep(1.0)
            if os.path.exists(done_flag):
                worker_finished = True
            break

        now = time.time()
        if now - last_message_time >= 10.0:
            print('Abaqus/CAE is open. Waiting for the pressure-sweep worker to finish...')
            last_message_time = now

        time.sleep(1.0)

    if process_return_code is None:
        process_return_code = process.poll()

    write_launch_diagnostics(
        run_root,
        launch_info,
        command,
        return_code=process_return_code
    )

    if not worker_finished:
        print('')
        print('Abaqus did not report a completed worker run.')
        print('Launch diagnostics: %s' % diagnostics_path)
        print('Worker error file (if created): %s' % error_file)
        return process_return_code if process_return_code not in (None, 0) else 1

    print('')
    print('Abaqus worker finished. Abaqus/CAE may remain open for you to inspect.')

    csv_path = os.path.join(run_root, CSV_NAME)
    results = read_results_csv(csv_path)

    if not results:
        print('No result rows were found. Check the Abaqus output in: %s' % jobs_dir)
        return 1

    graph_path = os.path.join(run_root, GRAPH_NAME)
    xlsx_path = os.path.join(run_root, XLSX_NAME)

    # path stuff
    model_label = base_name
    created_graph = create_graph(results, graph_path, model_label)
    created_excel = create_excel(results, xlsx_path, created_graph, model_label)

    print('')
    print('=' * 78)
    print('RUN FINISHED')
    print('=' * 78)
    print('Run folder: %s' % run_root)
    print('Abaqus job files: %s' % jobs_dir)
    print('CSV: %s' % csv_path)
    if created_excel:
        print('Excel: %s' % created_excel)
    if created_graph:
        print('Graph: %s' % created_graph)
    print('=' * 78)

    return 0


#
# abaqus stuff
####

#
# main part
#
def parse_last_step_time_lpf(sta_path):
    """
    Read the last STEP TIME/LPF value printed in an Abaqus/Standard .sta file.

    Typical status rows are arranged like:
    STEP INC ATT ... TOTAL_TIME/FREQ STEP_TIME/LPF INC_OF_TIME/LPF ...

    STEP TIME/LPF is therefore the eighth whitespace-separated value in a
    normal increment row (index 7). Rows such as 1U are also accepted because
    they still report the most recently reached step time.
    """
    if not os.path.exists(sta_path):
        return None

    last_value = None

    handle = open(sta_path, 'r')
    try:
        for line in handle:
            parts = line.split()
            if len(parts) < 9:
                continue

            # table part
            try:
                int(parts[0])
                int(parts[1])
            except:
                continue

            # data table
            # this part
            total_time = safe_float(parts[6])
            step_time_lpf = safe_float(parts[7])
            increment_value = safe_float(parts[8])

            if total_time is None or step_time_lpf is None or increment_value is None:
                continue

            last_value = step_time_lpf
    finally:
        handle.close()

    return last_value


def extract_odb_data(open_odb_function, odb_path, requested_step, requested_frame):
    odb = None

    try:
        odb = open_odb_function(path=odb_path, readOnly=True)

        if requested_step not in odb.steps:
            print("  ERROR: Step '%s' not found in ODB." % requested_step)
            return None

        step_obj = odb.steps[requested_step]
        frames = step_obj.frames

        if len(frames) == 0:
            print('  ERROR: No frames found in requested step.')
            return None

        if requested_frame == -1:
            frame = frames[-1]
        else:
            if requested_frame < 0 or requested_frame >= len(frames):
                print('  ERROR: Requested frame does not exist.')
                return None
            frame = frames[requested_frame]

        max_stress = None
        max_peeq = None

        if 'S' in frame.fieldOutputs:
            stress_values = []
            for value in frame.fieldOutputs['S'].values:
                try:
                    stress_values.append(value.mises)
                except:
                    pass
            if stress_values:
                max_stress = max(stress_values)

        if 'PEEQ' in frame.fieldOutputs:
            peeq_values = []
            for value in frame.fieldOutputs['PEEQ'].values:
                try:
                    peeq_values.append(value.data)
                except:
                    pass
            if peeq_values:
                max_peeq = max(peeq_values)

        # inputs here
        # get the inputs
        # input values
        frame_fraction = None
        try:
            step_period = float(step_obj.timePeriod)
            if step_period != 0:
                frame_fraction = float(frame.frameValue) / step_period
        except:
            pass

        return {
            'Max Mises Stress': max_stress,
            'Max PEEQ': max_peeq,
            'ODB Frame Fraction': frame_fraction
        }

    except Exception as exc:
        print('  ERROR reading ODB: %s' % exc)
        return None

    finally:
        if odb is not None:
            try:
                odb.close()
            except:
                pass


def write_worker_csv(results, csv_path):
    fieldnames = [
        'Attempted Pressure (MPa)',
        'Step Time/LPF',
        'Effective Pressure (MPa)',
        'Job Name',
        'Status',
        'Max Mises Stress',
        'Max PEEQ',
        'Failure Stop'
    ]

    # output stuff
    if sys.version_info[0] < 3:
        handle = open(csv_path, 'wb')
    else:
        handle = open(csv_path, 'w', newline='')

    try:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writerow(dict((name, name) for name in fieldnames))
        for row in results:
            writer.writerow(row)
    finally:
        handle.close()


####
# this part
####
def _safe_job_model_name(job):
    try:
        return str(job.model)
    except:
        return ''


def _find_existing_job_for_model(mdb, model_name, preferred_job_name=None):
    """Find the CAE job that is actually attached to model_name."""
    if preferred_job_name and preferred_job_name in mdb.jobs.keys():
        candidate = mdb.jobs[preferred_job_name]
        if _safe_job_model_name(candidate) == model_name:
            return candidate

    for name in mdb.jobs.keys():
        candidate = mdb.jobs[name]
        if _safe_job_model_name(candidate) == model_name:
            return candidate

    return None


def _archive_old_job_files(job_name, work_dir, archive_root):
    """
    Move old files for this exact job name out of the way before resubmitting.
    This prevents Abaqus from reading a stale ODB/STA from an earlier run.
    """
    if not os.path.isdir(archive_root):
        os.makedirs(archive_root)

    old_files = glob.glob(os.path.join(work_dir, job_name + '.*'))
    for old_path in old_files:
        if not os.path.isfile(old_path):
            continue
        base = os.path.basename(old_path)
        target = os.path.join(archive_root, base)
        if os.path.exists(target):
            root, ext = os.path.splitext(base)
            target = os.path.join(archive_root, root + '_' + timestamp_text() + ext)
        try:
            shutil.move(old_path, target)
        except Exception as exc:
            print('WARNING: Could not archive old job file %s: %s' % (old_path, exc))


#
# main part
#
def _copy_current_job_files(job_name, work_dir, jobs_dir, preexisting_input=False):
    """
    Collect the completed job into the new PressureSweep/Jobs folder.

    IMPORTANT: the Abaqus run itself is still performed from the ORIGINAL CAE
    directory exactly like working v4. Files are only collected/cleaned AFTER
    job completion and AFTER ODB/STA post-processing, so submission behavior is
    not changed.

    The .inp is copied to Jobs. It is removed from the source folder only when
    it did not exist before this run (normal model-based jobs regenerate it).
    Clearly generated solver/result files are removed from the source folder
    after a successful copy so the weldolet project folder is not left cluttered.
    """
    copied = []
    generated_extensions = set([
        '.odb', '.sta', '.msg', '.dat', '.log', '.com', '.prt', '.sim',
        '.mdl', '.pac', '.res', '.sel', '.stt', '.abq', '.023', '.lck'
    ])

    for source_path in glob.glob(os.path.join(work_dir, job_name + '.*')):
        if not os.path.isfile(source_path):
            continue

        target_path = os.path.join(jobs_dir, os.path.basename(source_path))
        copied_ok = False
        try:
            shutil.copy2(source_path, target_path)
            copied.append(target_path)
            copied_ok = True
        except Exception as exc:
            print('WARNING: Could not copy job file %s: %s' % (source_path, exc))

        if not copied_ok:
            continue

        ext = os.path.splitext(source_path)[1].lower()
        remove_from_source = ext in generated_extensions
        if ext == '.inp' and not preexisting_input:
            remove_from_source = True

        if remove_from_source:
            try:
                os.remove(source_path)
            except Exception as exc:
                print('WARNING: Could not clean generated file %s: %s' % (source_path, exc))

    return copied


def _tail_text_file(path, maximum_lines=120):
    if not path or not os.path.exists(path):
        return ''
    try:
        handle = open(path, 'r')
        try:
            lines = handle.readlines()
        finally:
            handle.close()
        return ''.join(lines[-maximum_lines:])
    except:
        return ''


def _write_failure_diagnostics(run_root, job_name, work_dir, status, submission_error=None):
    """Write the useful tail of Abaqus text outputs for the first failed job."""
    path = os.path.join(run_root, 'Failure_Diagnostics.txt')
    try:
        handle = open(path, 'w')
        try:
            handle.write('ABAQUS FIRST-FAILURE DIAGNOSTICS\n')
            handle.write('================================\n\n')
            handle.write('Job: %s\n' % job_name)
            handle.write('Status: %s\n' % status)
            handle.write('Working directory used for Abaqus: %s\n' % work_dir)
            if submission_error is not None:
                handle.write('Submission exception: %s\n' % str(submission_error))

            for ext in ('.dat', '.msg', '.sta', '.log'):
                source = os.path.join(work_dir, job_name + ext)
                tail = _tail_text_file(source)
                handle.write('\n\n===== %s =====\n' % ext)
                if tail:
                    handle.write(tail)
                else:
                    handle.write('[file missing or empty]\n')
        finally:
            handle.close()
    except Exception as exc:
        print('WARNING: Could not write failure diagnostics: %s' % exc)
    return path


####
# this part
####
def _create_job_from_template(mdb, job_name, model_name, base_job, work_dir):
    """
    Create a job only when the CAE does not already have one for this model.
    Existing model-specific jobs are always preferred because they preserve the
    user's exact original job setup.
    """
    if base_job is not None:
        kwargs = {
            'name': job_name,
            'model': model_name,
            'scratch': work_dir
        }
        # run the jobs
        for attr in ('type', 'numCpus', 'numDomains', 'memory', 'memoryUnits',
                     'multiprocessingMode', 'explicitPrecision',
                     'nodalOutputPrecision', 'resultsFormat', 'numGPUs'):
            try:
                kwargs[attr] = getattr(base_job, attr)
            except:
                pass
        try:
            return mdb.Job(**kwargs)
        except Exception as exc:
            print('Could not copy every template job option: %s' % exc)
            print('Creating the job with basic Abaqus defaults instead.')

    return mdb.Job(name=job_name, model=model_name, scratch=work_dir)


def _pressure_from_job_or_model(job_name, model_name):
    """Read pressure from names such as 6-MPa or Job-6-MPa3ur."""
    import re
    for text in (model_name, job_name):
        if not text:
            continue
        match = re.search(r'(^|[^0-9])([0-9]+(?:\.[0-9]+)?)\s*-\s*MPa', str(text), re.I)
        if match:
            try:
                return float(match.group(2))
            except:
                pass
    return None


def _copy_old_results_without_moving(job_name, work_dir, archive_root):
    """Back up existing job files but leave them in place so an existing job is not broken."""
    if not os.path.isdir(archive_root):
        os.makedirs(archive_root)

    for old_path in glob.glob(os.path.join(work_dir, job_name + '.*')):
        if not os.path.isfile(old_path):
            continue
        base = os.path.basename(old_path)
        target = os.path.join(archive_root, base)
        if os.path.exists(target):
            root, ext = os.path.splitext(base)
            target = os.path.join(archive_root, root + '_' + timestamp_text() + ext)
        try:
            shutil.copy2(old_path, target)
        except Exception as exc:
            print('WARNING: Could not back up old job file %s: %s' % (old_path, exc))


#
# main part
#
def _existing_pressure_jobs(mdb_object):
    """
    Return EXISTING CAE jobs in increasing pressure order.

    Nothing is created, copied, renamed, or re-pointed. This intentionally
    behaves like selecting the existing jobs in Job Manager and submitting them
    one at a time.
    """
    found = []
    for job_name in mdb_object.jobs.keys():
        job = mdb_object.jobs[job_name]
        model_name = _safe_job_model_name(job)
        pressure = _pressure_from_job_or_model(job_name, model_name)
        if pressure is None:
            continue
        if pressure < float(START_PRESSURE) - 1.0e-12:
            continue
        if pressure > float(END_PRESSURE) + 1.0e-12:
            continue
        found.append((pressure, str(job_name), str(model_name)))

    found.sort(key=lambda item: (item[0], item[1]))
    return found


def run_abaqus_worker():
    # abaqus run
    # save files
    # main part
    # abaqus part
    from abaqus import mdb
    from abaqusConstants import OFF
    from odbAccess import openOdb

    cae_path = os.environ.get(ENV_CAE_PATH)
    run_root = os.environ.get(ENV_RUN_ROOT)
    jobs_dir = os.environ.get(ENV_JOBS_DIR)
    source_dir = os.environ.get(ENV_SOURCE_DIR)

    if not run_root:
        run_root = os.getcwd()
    if not jobs_dir:
        jobs_dir = os.path.join(run_root, 'Jobs')
    if not source_dir or not os.path.isdir(source_dir):
        source_dir = os.path.dirname(os.path.abspath(cae_path)) if cae_path else os.getcwd()

    if not os.path.exists(jobs_dir):
        os.makedirs(jobs_dir)

    archive_dir = os.path.join(jobs_dir, 'Previous_Job_Files')
    csv_path = os.path.join(run_root, CSV_NAME)

    print('')
    print('=' * 78)
    print('ABAQUS-NATIVE PRESSURE SWEEP')
    print('=' * 78)
    try:
        print('CAE already open in Abaqus: %s' % str(mdb.pathName))
    except:
        print('CAE is already open in the current Abaqus session.')
    print('Job working directory (working v4 behavior): %s' % source_dir)
    print('Final collected job files: %s' % jobs_dir)
    print('MODE: submit EXISTING CAE jobs only; do not create or modify models/jobs.')
    print('=' * 78)

    pressure_jobs = _existing_pressure_jobs(mdb)
    if not pressure_jobs:
        raise Exception(
            'No existing jobs with pressure names like Job-6-MPa... or models like 6-MPa were found.'
        )

    print('Existing jobs that will be run:')
    for pressure, job_name, model_name in pressure_jobs:
        print('  %s MPa -> %s  [model: %s]' % (pressure_text(pressure), job_name, model_name))

    # abaqus run
    # save files
    # this part
    os.chdir(source_dir)

    results = []

    for pressure, job_name, model_name in pressure_jobs:
        job = mdb.jobs[job_name]
        p_text = pressure_text(pressure)

        print('')
        print('-' * 78)
        print('SUBMITTING EXISTING JOB: %s  (%s MPa)' % (job_name, p_text))
        print('Model already attached to job: %s' % model_name)
        print('-' * 78)

        # inputs here
        # main part
        preexisting_input = os.path.exists(os.path.join(source_dir, job_name + '.inp'))

        # stuff below
        # file stuff
        # inputs here
        _copy_old_results_without_moving(job_name, source_dir, archive_dir)

        # this part
        # jobs here
        # job stuff
        stale_lock = os.path.join(source_dir, job_name + '.lck')
        if os.path.exists(stale_lock):
            try:
                os.remove(stale_lock)
                print('Removed stale lock: %s' % stale_lock)
            except Exception as exc:
                print('WARNING: Could not remove stale lock: %s' % exc)

        status = 'UNKNOWN'
        submission_error = None

        try:
            # abaqus part
            # job part
            job.submit(consistencyChecking=OFF)
            job.waitForCompletion()
            status = str(job.status)
        except Exception as exc:
            submission_error = exc
            try:
                status = str(job.status)
            except:
                status = 'SUBMISSION ERROR'
            print('ERROR while submitting/running %s: %s' % (job_name, exc))

        print('Job status: %s' % status)

        odb_path = os.path.join(source_dir, job_name + '.odb')
        sta_path = os.path.join(source_dir, job_name + '.sta')
        completed = str(status).upper() == 'COMPLETED'
        failure_stop = 'NO'

        # model stuff
        # input stuff
        # abaqus part
        solver_started = os.path.exists(sta_path) and os.path.getsize(sta_path) > 0
        odb_data = None
        if os.path.exists(odb_path) and (completed or solver_started):
            odb_data = extract_odb_data(openOdb, odb_path, STEP_NAME, FRAME_NUMBER)
        elif os.path.exists(odb_path):
            print('Skipping ODB because the solver did not produce a usable status file.')
        else:
            print('ODB file was not found: %s' % odb_path)

        max_stress = odb_data.get('Max Mises Stress') if odb_data else None
        max_peeq = odb_data.get('Max PEEQ') if odb_data else None

        if completed:
            lpf_value = 1.0
            effective_pressure = pressure
        else:
            failure_stop = 'YES'
            lpf_value = parse_last_step_time_lpf(sta_path)

            if lpf_value is None and odb_data is not None:
                lpf_value = odb_data.get('ODB Frame Fraction')

            effective_pressure = pressure * lpf_value if lpf_value is not None else None
            diagnostic_path = _write_failure_diagnostics(
                run_root, job_name, source_dir, status, submission_error
            )

            print('')
            print('*** FIRST FAILED JOB - STOPPING HERE ***')
            print('Attempted pressure: %s MPa' % p_text)
            print('Job status: %s' % status)
            print('Last STEP TIME/LPF: %s' % str(lpf_value))
            print('Adjusted failure pressure: %s MPa' % str(effective_pressure))
            print('Failure diagnostics: %s' % diagnostic_path)

        result_row = {
            'Attempted Pressure (MPa)': pressure,
            'Step Time/LPF': lpf_value,
            'Effective Pressure (MPa)': effective_pressure,
            'Job Name': job_name,
            'Status': status,
            'Max Mises Stress': max_stress,
            'Max PEEQ': max_peeq,
            'Failure Stop': failure_stop
        }
        results.append(result_row)
        write_worker_csv(results, csv_path)

        copied_files = _copy_current_job_files(
            job_name, source_dir, jobs_dir, preexisting_input=preexisting_input
        )
        print('Max Mises stress: %s' % str(max_stress))
        print('Max PEEQ: %s' % str(max_peeq))
        print('CSV updated: %s' % csv_path)
        print('Collected %d current job files in: %s' % (len(copied_files), jobs_dir))

        # job part
        # jobs here
        if not completed:
            break

    # abaqus stuff
    # abaqus part
    try:
        mdb.save()
        print('Copied CAE saved after sweep.')
    except Exception as exc:
        print('WARNING: Could not save copied CAE: %s' % exc)

    print('')
    print('=' * 78)
    print('ABAQUS-NATIVE WORKER FINISHED')
    print('Result CSV: %s' % csv_path)
    print('Collected job files: %s' % jobs_dir)
    print('=' * 78)

    return 0


#
# keep this here
####

if __name__ == '__main__':
    if os.environ.get(ENV_WORKER_MODE) == '1':
        worker_root = os.environ.get(ENV_RUN_ROOT) or os.getcwd()
        done_path = os.path.join(worker_root, WORKER_DONE_NAME)
        error_path = os.path.join(worker_root, WORKER_ERROR_NAME)

        try:
            exit_code = run_abaqus_worker()
            handle = open(done_path, 'w')
            try:
                handle.write('Abaqus worker completed with exit code %s\n' % exit_code)
            finally:
                handle.close()
        except Exception as exc:
            import traceback
            try:
                handle = open(error_path, 'w')
                try:
                    handle.write('Abaqus worker failed.\n\n')
                    handle.write(str(exc) + '\n\n')
                    handle.write(traceback.format_exc())
                finally:
                    handle.close()
            except:
                pass
            raise
    else:
        exit_code = run_normal_python_launcher()
        sys.exit(exit_code)
