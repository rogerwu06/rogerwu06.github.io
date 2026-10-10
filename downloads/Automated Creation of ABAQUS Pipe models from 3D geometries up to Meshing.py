# -*- coding: utf-8 -*-
from __future__ import print_function

"""
STEP -> EXISTING ABAQUS CAE, SIMPLE 3-INPUT / IN-PLACE VERSION
================================================================

This is deliberately simpler than the previous launcher.

The ONLY user-selected paths are:
    1) Existing Abaqus .cae
    2) Replacement .step / .stp
    3) Abaqus itself (.lnk, abq2023.bat/.cmd, or Abaqus launcher)

There is NO output CAE input.
The selected CAE is updated in place, but a timestamped backup is created first.
The live log and transfer report are created automatically beside the CAE.

Normal Python launches the GUI.  The same file is then run by Abaqus Python in
noGUI mode.  Paths are passed through environment variables, so there are no
extra command-line path arguments to maintain.

IMPORTANT:
- Close the selected CAE in any other Abaqus session before running this.
- No partitions are created here.
- Automatic face/set mapping should still be visually checked after transfer.
"""

import os
import sys
import math
import re
import time
import shutil
import subprocess
import traceback
from datetime import datetime

WORKER_FLAG = 'STEP_TRANSFER_ABAQUS_WORKER'
ENV_CAE = 'STEP_TRANSFER_CAE'
ENV_STEP = 'STEP_TRANSFER_STEP'
ENV_LOG = 'STEP_TRANSFER_LIVE_LOG'
ENV_MODE = 'STEP_TRANSFER_MODE'

def _safe_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return default


def _distance3(a, b):
    return math.sqrt(
        (a[0] - b[0]) ** 2 +
        (a[1] - b[1]) ** 2 +
        (a[2] - b[2]) ** 2
    )


####
# this part
#
def _dot(a, b):
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]


def _norm(v):
    return math.sqrt(_dot(v, v))


def _unit(v):
    n = _norm(v)
    if n <= 1.0e-30:
        return (0.0, 0.0, 0.0)
    return (v[0]/n, v[1]/n, v[2]/n)


##
#
# #
def _bbox_dims(bb):
    low = bb['low']
    high = bb['high']
    return (
        max(0.0, high[0]-low[0]),
        max(0.0, high[1]-low[1]),
        max(0.0, high[2]-low[2])
    )


def _bbox_center(bb):
    low = bb['low']
    high = bb['high']
    return (
        0.5*(low[0]+high[0]),
        0.5*(low[1]+high[1]),
        0.5*(low[2]+high[2])
    )


def _bbox_diag(bb):
    d = _bbox_dims(bb)
    return max(math.sqrt(d[0]*d[0] + d[1]*d[1] + d[2]*d[2]), 1.0e-12)


#
# main part
####
def _merge_bboxes(boxes):
    if not boxes:
        return {'low': (0.0, 0.0, 0.0), 'high': (1.0, 1.0, 1.0)}

    lo = [1.0e300, 1.0e300, 1.0e300]
    hi = [-1.0e300, -1.0e300, -1.0e300]

    for bb in boxes:
        for i in range(3):
            lo[i] = min(lo[i], bb['low'][i])
            hi[i] = max(hi[i], bb['high'][i])

    return {'low': tuple(lo), 'high': tuple(hi)}


def _normalized_point(point, bb):
    low = bb['low']
    dims = _bbox_dims(bb)
    out = []
    for i in range(3):
        if dims[i] <= 1.0e-12:
            out.append(0.5)
        else:
            out.append((point[i]-low[i])/dims[i])
    return tuple(out)


def _shape_axis_ratios(bb):
    d = _bbox_dims(bb)
    m = max(max(d), 1.0e-12)
    return (d[0]/m, d[1]/m, d[2]/m)


# #
#
##
def _part_match_descriptor(bb, global_bb):
    return {
        'bbox': bb,
        'axis_ratio': _shape_axis_ratios(bb),
        'relative_center': _normalized_point(_bbox_center(bb), global_bb)
    }


def _part_match_score(a, b):
    # input stuff
    #
    axis_score = _distance3(a['axis_ratio'], b['axis_ratio'])

    # keep this
    # main part
    center_score = _distance3(a['relative_center'], b['relative_center'])

    return 3.0*axis_score + 2.0*center_score


def _resolve_shortcut(shortcut_path):
    """Resolve a Windows .lnk without executing it."""
    ps_path = os.path.abspath(shortcut_path).replace("'", "''")
    ps = (
        "$ErrorActionPreference='Stop';"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        "$w=New-Object -ComObject WScript.Shell;"
        "$s=$w.CreateShortcut('%s');"
        "Write-Output ('TARGET=' + $s.TargetPath);"
        "Write-Output ('ARGS=' + $s.Arguments);"
        "Write-Output ('WORKDIR=' + $s.WorkingDirectory);"
    ) % ps_path

    flags = 0
    try:
        flags = subprocess.CREATE_NO_WINDOW
    except Exception:
        pass

    out = subprocess.check_output(
        ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps],
        universal_newlines=True,
        creationflags=flags
    )
    info = {'target': '', 'args': '', 'workdir': ''}
    for raw in out.splitlines():
        line = raw.strip()
        if line.startswith('TARGET='):
            info['target'] = line[7:].strip()
        elif line.startswith('ARGS='):
            info['args'] = line[5:].strip()
        elif line.startswith('WORKDIR='):
            info['workdir'] = line[8:].strip()
    return info


####
# this part
#
def _extract_command_files(text):
    out = []
    if not text:
        return out
    pats = [
        r'"([^\"]*(?:abq|abaqus)[^\"]*\.(?:bat|cmd))"',
        r"'([^']*(?:abq|abaqus)[^']*\.(?:bat|cmd))'",
        r'([^\s\"]*(?:abq|abaqus)[^\s\"]*\.(?:bat|cmd))'
    ]
    for pat in pats:
        for m in re.finditer(pat, text, re.IGNORECASE):
            p = m.group(1).strip().strip('"').strip("'")
            if p and p not in out:
                out.append(p)
    return out


def _find_abaqus_command(selected_path):
    """
    The GUI has only one Abaqus field.  This resolves it internally.
    Returns a .bat/.cmd when possible, otherwise a directly selected executable.
    """
    p = os.path.abspath(selected_path.strip().strip('"'))
    if not os.path.isfile(p):
        raise RuntimeError('Abaqus path does not exist: %s' % p)

    candidates = []
    direct_exe = ''

    low = p.lower()
    if low.endswith('.lnk'):
        info = _resolve_shortcut(p)
        target = info.get('target', '')
        args = info.get('args', '')
        workdir = info.get('workdir', '')

        if target.lower().endswith(('.bat', '.cmd')):
            candidates.append(target)
        candidates.extend(_extract_command_files(target))
        candidates.extend(_extract_command_files(args))

        # calc part
        roots = []
        for seed in (target, workdir):
            if not seed:
                continue
            q = seed
            if os.path.isfile(q):
                q = os.path.dirname(q)
            for _ in range(8):
                if not q or q in roots:
                    break
                roots.append(q)
                parent = os.path.dirname(q)
                if parent == q:
                    break
                q = parent
        for root in roots:
            for rel in (
                os.path.join('Commands', 'abq2023.bat'),
                os.path.join('Commands', 'abaqus.bat'),
                os.path.join('commands', 'abq2023.bat'),
                'abq2023.bat',
                'abaqus.bat'
            ):
                candidates.append(os.path.join(root, rel))

        if target and os.path.isfile(target) and target.lower().endswith('.exe'):
            direct_exe = target

    elif low.endswith(('.bat', '.cmd')):
        candidates.append(p)
    elif low.endswith('.exe'):
        direct_exe = p
    else:
        candidates.append(p)

    #
    drive = os.environ.get('SystemDrive', 'C:')
    candidates.extend([
        os.path.join(drive + os.sep, 'SIMULIA', 'Commands', 'abq2023.bat'),
        os.path.join(drive + os.sep, 'SIMULIA', 'Commands', 'abaqus.bat'),
        os.path.join(drive + os.sep, 'SIMULIA', 'Abaqus', 'Commands', 'abq2023.bat'),
        os.path.join(drive + os.sep, 'SIMULIA', 'Abaqus', 'Commands', 'abaqus.bat')
    ])

    seen = set()
    for c in candidates:
        if not c:
            continue
        c = os.path.expandvars(c.strip().strip('"'))
        key = c.lower()
        if key in seen:
            continue
        seen.add(key)
        if os.path.isfile(c) and c.lower().endswith(('.bat', '.cmd')):
            return c

    if direct_exe:
        return direct_exe

    raise RuntimeError(
        'Could not locate the real Abaqus command from the selected path. '
        'Select the Abaqus CAE shortcut or abq2023.bat directly.'
    )


def _timestamped_backup(cae_path):
    folder = os.path.dirname(cae_path)
    stem, ext = os.path.splitext(os.path.basename(cae_path))
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = os.path.join(folder, stem + '_BEFORE_STEP_TRANSFER_' + stamp + ext)
    shutil.copy2(cae_path, backup)
    return backup


# #
# keep this
##
def _cae_lock_candidates(cae_path):
    base, _ = os.path.splitext(cae_path)
    return [cae_path + '.lck', base + '.lck']


def _build_abaqus_process(command_file, script_file):
    low = command_file.lower()
    worker_arg = 'noGUI=%s' % ('"' + script_file + '"')
    if low.endswith(('.bat', '.cmd')):
        inner = 'call "' + command_file + '" cae ' + worker_arg
        return ['cmd.exe', '/d', '/s', '/c', inner]
    return [command_file, 'cae', worker_arg]


def _normal_python_gui():
    try:
        try:
            import tkinter as tk
            from tkinter import filedialog, messagebox
            from tkinter.scrolledtext import ScrolledText
            import queue
        except ImportError:
            import Tkinter as tk
            import tkFileDialog as filedialog
            import tkMessageBox as messagebox
            from ScrolledText import ScrolledText
            import Queue as queue
        import threading
    except Exception as e:
        print('Could not start GUI: %s' % e)
        return 2

    root = tk.Tk()
    root.title('STEP -> Abaqus CAE | Probe-first 3-input transfer')
    root.geometry('980x650')
    root.minsize(820, 560)

    cae_var = tk.StringVar(value='')
    step_var = tk.StringVar(value='')
    abaqus_var = tk.StringVar(value='')
    status_var = tk.StringVar(value='Select the same 3 inputs. Full transfer now runs a minimal probe first.')

    q = queue.Queue()
    holder = {'proc': None, 'log_handle': None, 'backup': '', 'report': '', 'live_log': '', 'phase': '', 'settings': None}

    def add_log(text):
        log_box.configure(state='normal')
        log_box.insert('end', text)
        log_box.see('end')
        log_box.configure(state='disabled')

    def browse_cae():
        p = filedialog.askopenfilename(title='Select existing Abaqus CAE', filetypes=[('Abaqus CAE', '*.cae'), ('All files', '*.*')])
        if p:
            cae_var.set(p)

    def browse_step():
        p = filedialog.askopenfilename(title='Select replacement STEP', filetypes=[('STEP geometry', '*.step *.stp'), ('All files', '*.*')])
        if p:
            step_var.set(p)

    def browse_abaqus():
        p = filedialog.askopenfilename(
            title='Select Abaqus CAE shortcut or command',
            filetypes=[('Abaqus / shortcut', '*.lnk *.bat *.cmd *.exe'), ('All files', '*.*')]
        )
        if p:
            abaqus_var.set(p)

    def set_buttons(running):
        run_btn.configure(state=('disabled' if running else 'normal'))
        try:
            probe_btn.configure(state=('disabled' if running else 'normal'))
        except Exception:
            pass
        stop_btn.configure(state=('normal' if running else 'disabled'))

    def reader_thread(proc, log_handle):
        try:
            while True:
                line = proc.stdout.readline()
                if not line:
                    if proc.poll() is not None:
                        break
                    time.sleep(0.05)
                    continue
                try:
                    log_handle.write(line)
                    log_handle.flush()
                except Exception:
                    pass
                q.put(('line', line))
            rc = proc.wait()
            q.put(('done', rc))
        except Exception as e:
            q.put(('error', str(e)))

    def _tail_text(path, max_chars=9000):
        try:
            with open(path, 'r') as f:
                text = f.read()
            if len(text) > max_chars:
                text = '... [earlier log omitted] ...\n' + text[-max_chars:]
            return text
        except Exception as e:
            return 'Could not read live log: %s' % e

    def _show_failure_details(return_code):
        live_log = holder.get('live_log', '')
        text = _tail_text(live_log)
        err_path = ''
        try:
            if live_log:
                err_path = os.path.splitext(live_log)[0] + '_ERROR.txt'
                with open(err_path, 'w') as f:
                    f.write('Abaqus return code: %s\n\n' % return_code)
                    f.write(text)
        except Exception:
            err_path = ''

        win = tk.Toplevel(root)
        win.title('Abaqus failure details')
        win.geometry('980x650')
        tk.Label(
            win,
            text='Abaqus returned code %s. The actual Abaqus message is shown below.' % return_code,
            anchor='w', justify='left', font=('Arial', 10, 'bold')
        ).pack(fill='x', padx=10, pady=(10, 4))
        if holder.get('backup'):
            tk.Label(
                win,
                text='Backup: %s' % holder.get('backup'),
                anchor='w', justify='left', wraplength=940
            ).pack(fill='x', padx=10, pady=2)
        tk.Label(
            win,
            text='Live log: %s%s' % (live_log, ('\nError copy: ' + err_path) if err_path else ''),
            anchor='w', justify='left', wraplength=940
        ).pack(fill='x', padx=10, pady=(2, 6))
        box = ScrolledText(win, wrap='word', font=('Consolas', 9))
        box.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        box.insert('1.0', text)
        box.configure(state='disabled')

    def _validated_inputs():
        cae = os.path.abspath(cae_var.get().strip().strip('"')) if cae_var.get().strip() else ''
        step = os.path.abspath(step_var.get().strip().strip('"')) if step_var.get().strip() else ''
        selected_abaqus = abaqus_var.get().strip().strip('"')

        problems = []
        if not cae or not os.path.isfile(cae):
            problems.append('Select a valid .cae file.')
        if not step or not os.path.isfile(step):
            problems.append('Select a valid .step/.stp file.')
        if not selected_abaqus or not os.path.isfile(selected_abaqus):
            problems.append('Select the Abaqus shortcut/command file.')
        if problems:
            messagebox.showerror('Missing input', '\n'.join(problems))
            return None

        for lock in _cae_lock_candidates(cae):
            if os.path.exists(lock):
                messagebox.showerror(
                    'CAE appears to be open',
                    'Abaqus lock file exists:\n%s\n\nClose this CAE in Abaqus before running.' % lock
                )
                return None

        try:
            command_file = _find_abaqus_command(selected_abaqus)
        except Exception as e:
            messagebox.showerror('Abaqus command problem', str(e))
            return None

        return cae, step, command_file

    def _launch_worker(cae, step, command_file, mode, backup=''):
        base = os.path.splitext(cae)[0]
        suffix = '_STEP_IMPORT_PROBE.log' if mode == 'probe' else '_STEP_TRANSFER_LIVE.log'
        live_log = base + suffix
        report = base + '_STEP_TRANSFER_REPORT.txt'

        env = os.environ.copy()
        env[WORKER_FLAG] = '1'
        env[ENV_CAE] = cae
        env[ENV_STEP] = step
        env[ENV_LOG] = live_log
        env[ENV_MODE] = mode

        script_file = os.path.abspath(__file__)
        args = _build_abaqus_process(command_file, script_file)

        add_log('\n' + '=' * 72 + '\n')
        add_log(('MINIMAL ABAQUS/STEP PROBE' if mode == 'probe' else 'FULL IN-PLACE STEP TRANSFER') + '\n')
        add_log('=' * 72 + '\n')
        add_log('CAE:    %s\n' % cae)
        add_log('STEP:   %s\n' % step)
        add_log('Abaqus: %s\n' % command_file)
        if backup:
            add_log('Backup: %s\n' % backup)
        add_log('Log:    %s\n\n' % live_log)

        try:
            log_handle = open(live_log, 'w')
            proc = subprocess.Popen(
                args,
                cwd=os.path.dirname(cae) or None,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1
            )
        except Exception as e:
            try:
                log_handle.close()
            except Exception:
                pass
            messagebox.showerror('Launch failed', str(e))
            return False

        holder['proc'] = proc
        holder['log_handle'] = log_handle
        holder['backup'] = backup
        holder['report'] = report
        holder['live_log'] = live_log
        holder['phase'] = mode
        holder['settings'] = (cae, step, command_file)
        status_var.set('Running minimal probe...' if mode == 'probe' else 'Abaqus is performing the full transfer...')
        set_buttons(True)

        t = threading.Thread(target=reader_thread, args=(proc, log_handle))
        t.daemon = True
        t.start()
        return True

    def run_probe_only():
        settings = _validated_inputs()
        if settings is None:
            return
        cae, step, command_file = settings
        _launch_worker(cae, step, command_file, 'probe', '')

    def run_transfer():
        settings = _validated_inputs()
        if settings is None:
            return
        cae, step, command_file = settings

        # needed here
        # this part
        holder['settings'] = settings
        holder['backup'] = ''
        holder['phase'] = 'probe_then_full'
        if _launch_worker(cae, step, command_file, 'probe', ''):
            holder['phase'] = 'probe_then_full'

    def stop_transfer():
        proc = holder.get('proc')
        if proc is not None and proc.poll() is None:
            try:
                proc.terminate()
                add_log('\nSTOP REQUESTED. The backup remains untouched.\n')
                status_var.set('Stop requested.')
            except Exception as e:
                messagebox.showerror('Stop failed', str(e))

    def poll():
        try:
            while True:
                kind, value = q.get_nowait()
                if kind == 'line':
                    add_log(value)
                elif kind == 'error':
                    add_log('\nOUTPUT READER ERROR: %s\n' % value)
                elif kind == 'done':
                    try:
                        if holder.get('log_handle'):
                            holder['log_handle'].close()
                    except Exception:
                        pass

                    phase = holder.get('phase', '')
                    settings = holder.get('settings')

                    # leave this
                    # calc part
                    if phase == 'probe_then_full':
                        if value == 0 and settings:
                            cae, step, command_file = settings
                            add_log('\nPROBE PASSED. CAE opened, STEP opened, and every STEP body imported in memory.\n')
                            try:
                                backup = _timestamped_backup(cae)
                            except Exception as e:
                                set_buttons(False)
                                status_var.set('Probe passed but backup creation failed.')
                                messagebox.showerror('Backup failed', str(e))
                                continue
                            holder['backup'] = backup
                            holder['phase'] = 'full'
                            _launch_worker(cae, step, command_file, 'full', backup)
                            holder['phase'] = 'full'
                            continue
                        else:
                            set_buttons(False)
                            status_var.set('Minimal probe failed. Full transfer was NOT started.')
                            _show_failure_details(value)
                            continue

                    if phase == 'probe':
                        set_buttons(False)
                        if value == 0:
                            status_var.set('Probe passed. CAE + STEP import are working.')
                            messagebox.showinfo(
                                'Probe passed',
                                'Abaqus successfully opened the CAE, opened the STEP file, and imported every STEP body in memory.\n\nNo CAE changes were saved.'
                            )
                        else:
                            status_var.set('Probe failed. See failure details.')
                            _show_failure_details(value)
                        continue

                    # input stuff
                    set_buttons(False)
                    if value == 0:
                        status_var.set('Transfer finished successfully.')
                        messagebox.showinfo(
                            'Transfer finished',
                            'The CAE was updated successfully.\n\nBackup:\n%s\n\nReport:\n%s' % (
                                holder.get('backup', ''), holder.get('report', '')
                            )
                        )
                    else:
                        status_var.set('Transfer failed. Backup is untouched.')
                        _show_failure_details(value)
        except queue.Empty:
            pass
        root.after(100, poll)

    outer = tk.Frame(root, padx=14, pady=12)
    outer.pack(fill='both', expand=True)

    tk.Label(outer, text='STEP -> Existing Abaqus CAE', font=('Arial', 16, 'bold'), anchor='w').pack(fill='x')
    tk.Label(
        outer,
        text='Only 3 inputs. The CAE is backed up automatically and updated in place. This version avoids scanning every unused internal Abaqus region and caches geometry matching.',
        anchor='w', justify='left', wraplength=930
    ).pack(fill='x', pady=(4, 10))

    form = tk.Frame(outer)
    form.pack(fill='x')
    form.columnconfigure(1, weight=1)

    def row(r, label, var, command):
        tk.Label(form, text=label, width=22, anchor='w').grid(row=r, column=0, sticky='w', pady=5)
        tk.Entry(form, textvariable=var).grid(row=r, column=1, sticky='ew', pady=5)
        tk.Button(form, text='Browse...', width=11, command=command).grid(row=r, column=2, padx=(8, 0), pady=5)

    row(0, 'Existing CAE', cae_var, browse_cae)
    row(1, 'Replacement STEP', step_var, browse_step)
    row(2, 'Abaqus location', abaqus_var, browse_abaqus)

    buttons = tk.Frame(outer)
    buttons.pack(fill='x', pady=(10, 8))
    probe_btn = tk.Button(buttons, text='TEST CAE + STEP ONLY', width=22, command=run_probe_only)
    probe_btn.pack(side='left', padx=(0, 8))
    run_btn = tk.Button(buttons, text='RUN FULL TRANSFER', width=20, command=run_transfer, font=('Arial', 10, 'bold'))
    run_btn.pack(side='left', padx=(0, 8))
    stop_btn = tk.Button(buttons, text='STOP', width=10, command=stop_transfer, state='disabled')
    stop_btn.pack(side='left')

    tk.Label(outer, textvariable=status_var, relief='sunken', anchor='w', padx=7, pady=5).pack(fill='x', pady=(0, 8))
    tk.Label(outer, text='Live Abaqus output:', anchor='w', font=('Arial', 10, 'bold')).pack(fill='x')
    log_box = ScrolledText(outer, height=24, wrap='word', font=('Consolas', 9))
    log_box.pack(fill='both', expand=True, pady=(4, 0))
    log_box.configure(state='disabled')

    root.after(100, poll)
    root.mainloop()
    return 0

# #
# keep this
##
def _abaqus_worker(template_cae, step_file):
    from abaqus import mdb, openMdb
    from abaqusConstants import (
        THREE_D,
        DEFORMABLE_BODY,
        ON,
        OFF
    )

    report_lines = []
    warnings = []

    #
    # this part
    #
    descriptor_cache = {}

    def log(text=''):
        text = str(text)
        print(text)
        try:
            sys.stdout.flush()
        except Exception:
            pass
        report_lines.append(text)

    def warn(text):
        text = 'WARNING: ' + str(text)
        print(text)
        try:
            sys.stdout.flush()
        except Exception:
            pass
        report_lines.append(text)
        warnings.append(text)

    def repo_keys(repo):
        try:
            return list(repo.keys())
        except Exception:
            return []

    def instance_bbox(inst):
        # calc part
        try:
            if len(inst.faces) > 0:
                return inst.faces.getBoundingBox()
        except Exception:
            pass

        try:
            if len(inst.cells) > 0:
                return inst.cells.getBoundingBox()
        except Exception:
            pass

        try:
            if len(inst.edges) > 0:
                return inst.edges.getBoundingBox()
        except Exception:
            pass

        raise Exception('Could not calculate bounding box for instance %s' % inst.name)

    def face_desc(face, owner_bb):
        try:
            c = tuple(face.getCentroid())
        except Exception:
            try:
                c = tuple(face.pointOn[0])
            except Exception:
                c = (0.0, 0.0, 0.0)

        try:
            n = tuple(face.getNormal(point=c))
        except Exception:
            try:
                n = tuple(face.getNormal())
            except Exception:
                n = (0.0, 0.0, 0.0)

        try:
            area = float(face.getSize(printResults=False))
        except Exception:
            area = 0.0

        try:
            edge_count = len(face.getEdges())
        except Exception:
            edge_count = 0

        inst_name = ''
        try:
            inst_name = str(face.instanceName)
        except Exception:
            pass

        return {
            'index': int(face.index),
            'point': c,
            'normal': _unit(n),
            'measure': area,
            'edge_count': edge_count,
            'normalized_point': _normalized_point(c, owner_bb),
            'instance_name': inst_name
        }

    def edge_desc(edge, owner_bb):
        try:
            p = tuple(edge.pointOn[0])
        except Exception:
            p = (0.0, 0.0, 0.0)

        try:
            length = float(edge.getSize(printResults=False))
        except Exception:
            length = 0.0

        inst_name = ''
        try:
            inst_name = str(edge.instanceName)
        except Exception:
            pass

        return {
            'index': int(edge.index),
            'point': p,
            'measure': length,
            'normalized_point': _normalized_point(p, owner_bb),
            'instance_name': inst_name
        }

    def vertex_desc(vertex, owner_bb):
        try:
            p = tuple(vertex.pointOn[0])
        except Exception:
            try:
                p = tuple(vertex.pointOn)
            except Exception:
                p = (0.0, 0.0, 0.0)

        inst_name = ''
        try:
            inst_name = str(vertex.instanceName)
        except Exception:
            pass

        return {
            'index': int(vertex.index),
            'point': p,
            'normalized_point': _normalized_point(p, owner_bb),
            'instance_name': inst_name
        }

    def cell_desc(cell, owner_bb):
        try:
            p = tuple(cell.pointOn[0])
        except Exception:
            p = _bbox_center(owner_bb)

        try:
            volume = float(cell.getSize(printResults=False))
        except Exception:
            volume = 0.0

        inst_name = ''
        try:
            inst_name = str(cell.instanceName)
        except Exception:
            pass

        return {
            'index': int(cell.index),
            'point': p,
            'measure': volume,
            'normalized_point': _normalized_point(p, owner_bb),
            'instance_name': inst_name
        }

    def entity_score(old_d, new_d, old_bb, new_bb, is_face=False):
        # input stuff
        s = 4.0 * _distance3(
            old_d['normalized_point'],
            new_d['normalized_point']
        )

        old_scale = _bbox_diag(old_bb)
        new_scale = _bbox_diag(new_bb)

        old_measure = max(_safe_float(old_d.get('measure', 0.0)), 0.0)
        new_measure = max(_safe_float(new_d.get('measure', 0.0)), 0.0)

        if old_measure > 0.0 and new_measure > 0.0:
            if is_face:
                old_m = old_measure / max(old_scale*old_scale, 1.0e-20)
                new_m = new_measure / max(new_scale*new_scale, 1.0e-20)
            else:
                old_m = old_measure / max(old_scale, 1.0e-20)
                new_m = new_measure / max(new_scale, 1.0e-20)

            ratio = max(old_m, new_m) / max(min(old_m, new_m), 1.0e-20)
            s += 0.35 * abs(math.log(max(ratio, 1.0)))

        if is_face:
            n1 = old_d.get('normal', (0.0, 0.0, 0.0))
            n2 = new_d.get('normal', (0.0, 0.0, 0.0))
            if _norm(n1) > 0.5 and _norm(n2) > 0.5:
                # output stuff
                s += 2.0 * (1.0 - abs(_dot(n1, n2)))

            e1 = int(old_d.get('edge_count', 0))
            e2 = int(new_d.get('edge_count', 0))
            if max(e1, e2) > 0:
                s += 0.15 * abs(e1-e2) / float(max(e1, e2))

        return s

    def _sequence_cache_key(seq, kind, owner_bb):
        inst_name = ''
        try:
            if len(seq) > 0:
                inst_name = str(seq[0].instanceName)
        except Exception:
            pass

        # keep this
        if inst_name:
            return (kind, inst_name, len(seq))

        # main part
        try:
            c = _bbox_center(owner_bb)
            d = _bbox_dims(owner_bb)
            return (
                kind,
                len(seq),
                round(c[0], 8), round(c[1], 8), round(c[2], 8),
                round(d[0], 8), round(d[1], 8), round(d[2], 8)
            )
        except Exception:
            return (kind, id(seq))

    def _cached_entities(seq, kind, owner_bb):
        key = _sequence_cache_key(seq, kind, owner_bb)
        if key in descriptor_cache:
            return descriptor_cache[key]

        vals = []
        if kind == 'face':
            for ent in seq:
                try:
                    vals.append((ent, face_desc(ent, owner_bb)))
                except Exception:
                    pass
        elif kind == 'edge':
            for ent in seq:
                try:
                    vals.append((ent, edge_desc(ent, owner_bb)))
                except Exception:
                    pass
        elif kind == 'vertex':
            for ent in seq:
                try:
                    vals.append((ent, vertex_desc(ent, owner_bb)))
                except Exception:
                    pass
        elif kind == 'cell':
            for ent in seq:
                try:
                    vals.append((ent, cell_desc(ent, owner_bb)))
                except Exception:
                    pass

        descriptor_cache[key] = vals
        return vals

    def _best_face(old_d, target_faces, old_bb, new_bb):
        best = None
        best_score = 1.0e300
        for f, nd in _cached_entities(target_faces, 'face', new_bb):
            try:
                sc = entity_score(old_d, nd, old_bb, new_bb, is_face=True)
                if sc < best_score:
                    best = f
                    best_score = sc
            except Exception:
                pass
        return best, best_score

    def _best_edge(old_d, target_edges, old_bb, new_bb):
        best = None
        best_score = 1.0e300
        for e, nd in _cached_entities(target_edges, 'edge', new_bb):
            try:
                sc = entity_score(old_d, nd, old_bb, new_bb, is_face=False)
                if sc < best_score:
                    best = e
                    best_score = sc
            except Exception:
                pass
        return best, best_score

    def _best_vertex(old_d, target_vertices, old_bb, new_bb):
        best = None
        best_score = 1.0e300
        for v, nd in _cached_entities(target_vertices, 'vertex', new_bb):
            try:
                sc = 4.0 * _distance3(
                    old_d['normalized_point'],
                    nd['normalized_point']
                )
                if sc < best_score:
                    best = v
                    best_score = sc
            except Exception:
                pass
        return best, best_score

    def _best_cell(old_d, target_cells, old_bb, new_bb):
        best = None
        best_score = 1.0e300
        for c, nd in _cached_entities(target_cells, 'cell', new_bb):
            try:
                sc = entity_score(old_d, nd, old_bb, new_bb, is_face=False)
                if sc < best_score:
                    best = c
                    best_score = sc
            except Exception:
                pass
        return best, best_score

    def _dedupe_by_index(items):
        out = []
        seen = set()
        for item in items:
            try:
                key = int(item.index)
            except Exception:
                key = id(item)
            if key not in seen:
                seen.add(key)
                out.append(item)
        return out

    def capture_set(set_obj, bbox_lookup, default_bbox):
        data = {
            'type': None,
            'items': []
        }

        # needed here
        # this part
        try:
            if len(set_obj.faces) > 0:
                data['type'] = 'faces'
                for x in set_obj.faces:
                    inst = ''
                    try:
                        inst = str(x.instanceName)
                    except Exception:
                        pass
                    bb = bbox_lookup.get(inst, default_bbox)
                    data['items'].append(face_desc(x, bb))
                return data
        except Exception:
            pass

        try:
            if len(set_obj.cells) > 0:
                data['type'] = 'cells'
                for x in set_obj.cells:
                    inst = ''
                    try:
                        inst = str(x.instanceName)
                    except Exception:
                        pass
                    bb = bbox_lookup.get(inst, default_bbox)
                    data['items'].append(cell_desc(x, bb))
                return data
        except Exception:
            pass

        try:
            if len(set_obj.edges) > 0:
                data['type'] = 'edges'
                for x in set_obj.edges:
                    inst = ''
                    try:
                        inst = str(x.instanceName)
                    except Exception:
                        pass
                    bb = bbox_lookup.get(inst, default_bbox)
                    data['items'].append(edge_desc(x, bb))
                return data
        except Exception:
            pass

        try:
            if len(set_obj.vertices) > 0:
                data['type'] = 'vertices'
                for x in set_obj.vertices:
                    inst = ''
                    try:
                        inst = str(x.instanceName)
                    except Exception:
                        pass
                    bb = bbox_lookup.get(inst, default_bbox)
                    data['items'].append(vertex_desc(x, bb))
                return data
        except Exception:
            pass

        #
        # calc part
        return data

    def capture_surface(surface_obj, bbox_lookup, default_bbox):
        data = {
            'sides': {},
            'sidedness_known': False
        }

        # input stuff
        for side_name in ('side1Faces', 'side2Faces', 'side12Faces'):
            try:
                seq = getattr(surface_obj, side_name)
                if seq is not None and len(seq) > 0:
                    vals = []
                    for x in seq:
                        inst = ''
                        try:
                            inst = str(x.instanceName)
                        except Exception:
                            pass
                        bb = bbox_lookup.get(inst, default_bbox)
                        vals.append(face_desc(x, bb))
                    if vals:
                        data['sides'][side_name] = vals
                        data['sidedness_known'] = True
            except Exception:
                pass

        if not data['sides']:
            # output stuff
            # keep this
            # main part
            try:
                vals = []
                for x in surface_obj.faces:
                    inst = ''
                    try:
                        inst = str(x.instanceName)
                    except Exception:
                        pass
                    bb = bbox_lookup.get(inst, default_bbox)
                    vals.append(face_desc(x, bb))
                if vals:
                    data['sides']['side1Faces'] = vals
            except Exception:
                pass

        return data

    def capture_repository_regions(owner, bbox_lookup, default_bbox):
        sets = {}
        surfaces = {}

        # needed here
        try:
            for name in repo_keys(owner.sets):
                try:
                    sets[name] = capture_set(owner.sets[name], bbox_lookup, default_bbox)
                except Exception:
                    pass
        except Exception:
            pass

        # this part
        try:
            for name in repo_keys(owner.allInternalSets):
                if name not in sets:
                    try:
                        sets[name] = capture_set(
                            owner.allInternalSets[name],
                            bbox_lookup,
                            default_bbox
                        )
                    except Exception:
                        pass
        except Exception:
            pass

        #
        try:
            for name in repo_keys(owner.surfaces):
                try:
                    surfaces[name] = capture_surface(
                        owner.surfaces[name],
                        bbox_lookup,
                        default_bbox
                    )
                except Exception:
                    pass
        except Exception:
            pass

        # calc part
        try:
            for name in repo_keys(owner.allInternalSurfaces):
                if name not in surfaces:
                    try:
                        surfaces[name] = capture_surface(
                            owner.allInternalSurfaces[name],
                            bbox_lookup,
                            default_bbox
                        )
                    except Exception:
                        pass
        except Exception:
            pass

        return sets, surfaces

    def _try_delete_named_region(owner, name, is_surface):
        if is_surface:
            repos = ('surfaces', 'allInternalSurfaces')
        else:
            repos = ('sets', 'allInternalSets')

        for rname in repos:
            try:
                repo = getattr(owner, rname)
                if name in repo_keys(repo):
                    del repo[name]
            except Exception:
                pass

    def map_part_set_to_new_part(
        name,
        captured,
        old_bb,
        new_temp_inst,
        new_part,
        new_bb
    ):
        typ = captured.get('type')
        old_items = captured.get('items', [])

        if not typ or not old_items:
            return False, 0.0, 0

        mapped = []
        scores = []

        if typ == 'faces':
            for d in old_items:
                ent, sc = _best_face(d, new_temp_inst.faces, old_bb, new_bb)
                if ent is not None:
                    mapped.append(new_part.faces[ent.index])
                    scores.append(sc)
            mapped = _dedupe_by_index(mapped)
            if mapped:
                try:
                    _try_delete_named_region(new_part, name, False)
                    new_part.Set(name=name, faces=tuple(mapped))
                    return True, max(scores) if scores else 0.0, len(mapped)
                except Exception as e:
                    warn('Could not recreate part set "%s": %s' % (name, e))

        elif typ == 'cells':
            for d in old_items:
                ent, sc = _best_cell(d, new_temp_inst.cells, old_bb, new_bb)
                if ent is not None:
                    mapped.append(new_part.cells[ent.index])
                    scores.append(sc)
            mapped = _dedupe_by_index(mapped)
            if mapped:
                try:
                    _try_delete_named_region(new_part, name, False)
                    new_part.Set(name=name, cells=tuple(mapped))
                    return True, max(scores) if scores else 0.0, len(mapped)
                except Exception as e:
                    warn('Could not recreate part cell set "%s": %s' % (name, e))

        elif typ == 'edges':
            for d in old_items:
                ent, sc = _best_edge(d, new_temp_inst.edges, old_bb, new_bb)
                if ent is not None:
                    mapped.append(new_part.edges[ent.index])
                    scores.append(sc)
            mapped = _dedupe_by_index(mapped)
            if mapped:
                try:
                    _try_delete_named_region(new_part, name, False)
                    new_part.Set(name=name, edges=tuple(mapped))
                    return True, max(scores) if scores else 0.0, len(mapped)
                except Exception as e:
                    warn('Could not recreate part edge set "%s": %s' % (name, e))

        elif typ == 'vertices':
            for d in old_items:
                ent, sc = _best_vertex(d, new_temp_inst.vertices, old_bb, new_bb)
                if ent is not None:
                    mapped.append(new_part.vertices[ent.index])
                    scores.append(sc)
            mapped = _dedupe_by_index(mapped)
            if mapped:
                try:
                    _try_delete_named_region(new_part, name, False)
                    new_part.Set(name=name, vertices=tuple(mapped))
                    return True, max(scores) if scores else 0.0, len(mapped)
                except Exception as e:
                    warn('Could not recreate part vertex set "%s": %s' % (name, e))

        return False, max(scores) if scores else 0.0, len(mapped)

    def map_part_surface_to_new_part(
        name,
        captured,
        old_bb,
        new_temp_inst,
        new_part,
        new_bb
    ):
        kwargs = {}
        scores = []

        for side_name, old_items in captured.get('sides', {}).items():
            mapped = []
            side_dots = []

            for d in old_items:
                ent, sc = _best_face(d, new_temp_inst.faces, old_bb, new_bb)
                if ent is not None:
                    mapped.append(new_part.faces[ent.index])
                    scores.append(sc)

                    try:
                        nd = face_desc(new_temp_inst.faces[ent.index], new_bb)
                        side_dots.append(_dot(d['normal'], nd['normal']))
                    except Exception:
                        pass

            mapped = _dedupe_by_index(mapped)

            if mapped:
                actual_side = side_name

                # input stuff
                #
                # keep this
                if captured.get('sidedness_known') and side_dots:
                    avg_dot = sum(side_dots) / float(len(side_dots))
                    if avg_dot < -0.5:
                        if side_name == 'side1Faces':
                            actual_side = 'side2Faces'
                        elif side_name == 'side2Faces':
                            actual_side = 'side1Faces'

                kwargs[actual_side] = tuple(mapped)

        if kwargs:
            try:
                _try_delete_named_region(new_part, name, True)
                new_part.Surface(name=name, **kwargs)
                return True, max(scores) if scores else 0.0, sum(len(x) for x in kwargs.values())
            except Exception as e:
                warn('Could not recreate part surface "%s": %s' % (name, e))

        return False, max(scores) if scores else 0.0, 0

    def assembly_target_instance_name(old_instance_name, instance_names):
        if old_instance_name in instance_names:
            return old_instance_name

        # main part
        #
        # this part
        if (not old_instance_name) and len(instance_names) == 1:
            return instance_names[0]

        return None

    def map_assembly_set(name, captured, assembly, old_bbox_lookup, new_bbox_lookup):
        typ = captured.get('type')
        old_items = captured.get('items', [])
        if not typ or not old_items:
            return False, 0.0, 0

        mapped = []
        scores = []
        instance_names = repo_keys(assembly.instances)

        for d in old_items:
            old_inst = d.get('instance_name', '')
            target_name = assembly_target_instance_name(old_inst, instance_names)
            if target_name is None:
                continue

            try:
                target = assembly.instances[target_name]
                old_bb = old_bbox_lookup.get(old_inst, old_bbox_lookup.get(target_name))
                new_bb = new_bbox_lookup.get(target_name)
                if old_bb is None or new_bb is None:
                    continue

                if typ == 'faces':
                    ent, sc = _best_face(d, target.faces, old_bb, new_bb)
                elif typ == 'cells':
                    ent, sc = _best_cell(d, target.cells, old_bb, new_bb)
                elif typ == 'edges':
                    ent, sc = _best_edge(d, target.edges, old_bb, new_bb)
                elif typ == 'vertices':
                    ent, sc = _best_vertex(d, target.vertices, old_bb, new_bb)
                else:
                    ent, sc = None, 0.0

                if ent is not None:
                    mapped.append(ent)
                    scores.append(sc)

            except Exception:
                pass

        mapped = _dedupe_by_index(mapped)

        if not mapped:
            return False, max(scores) if scores else 0.0, 0

        try:
            _try_delete_named_region(assembly, name, False)

            if typ == 'faces':
                assembly.Set(name=name, faces=tuple(mapped))
            elif typ == 'cells':
                assembly.Set(name=name, cells=tuple(mapped))
            elif typ == 'edges':
                assembly.Set(name=name, edges=tuple(mapped))
            elif typ == 'vertices':
                assembly.Set(name=name, vertices=tuple(mapped))
            else:
                return False, max(scores) if scores else 0.0, 0

            return True, max(scores) if scores else 0.0, len(mapped)

        except Exception as e:
            # leave this
            # calc part
            # input stuff
            warn('Could not recreate assembly set "%s": %s' % (name, e))
            return False, max(scores) if scores else 0.0, len(mapped)

    def map_assembly_surface(name, captured, assembly, old_bbox_lookup, new_bbox_lookup):
        kwargs = {}
        scores = []
        instance_names = repo_keys(assembly.instances)

        for side_name, old_items in captured.get('sides', {}).items():
            mapped = []
            dots = []

            for d in old_items:
                old_inst = d.get('instance_name', '')
                target_name = assembly_target_instance_name(old_inst, instance_names)
                if target_name is None:
                    continue

                try:
                    target = assembly.instances[target_name]
                    old_bb = old_bbox_lookup.get(old_inst, old_bbox_lookup.get(target_name))
                    new_bb = new_bbox_lookup.get(target_name)
                    if old_bb is None or new_bb is None:
                        continue

                    ent, sc = _best_face(d, target.faces, old_bb, new_bb)
                    if ent is not None:
                        mapped.append(ent)
                        scores.append(sc)
                        try:
                            nd = face_desc(ent, new_bb)
                            dots.append(_dot(d['normal'], nd['normal']))
                        except Exception:
                            pass
                except Exception:
                    pass

            mapped = _dedupe_by_index(mapped)

            if mapped:
                actual_side = side_name

                if captured.get('sidedness_known') and dots:
                    avg_dot = sum(dots) / float(len(dots))
                    if avg_dot < -0.5:
                        if side_name == 'side1Faces':
                            actual_side = 'side2Faces'
                        elif side_name == 'side2Faces':
                            actual_side = 'side1Faces'

                kwargs[actual_side] = tuple(mapped)

        if not kwargs:
            return False, max(scores) if scores else 0.0, 0

        try:
            _try_delete_named_region(assembly, name, True)
            assembly.Surface(name=name, **kwargs)
            return True, max(scores) if scores else 0.0, sum(len(x) for x in kwargs.values())
        except Exception as e:
            warn('Could not recreate assembly surface "%s": %s' % (name, e))
            return False, max(scores) if scores else 0.0, sum(len(x) for x in kwargs.values())

    def attr_region_name(obj):
        try:
            region_tuple = obj.region
            if region_tuple and len(region_tuple) > 0:
                return str(region_tuple[0])
        except Exception:
            pass
        return None

    def named_region_exists(model, region_name):
        if not region_name:
            return False

        a = model.rootAssembly

        for repo_name in ('sets', 'allInternalSets', 'surfaces', 'allInternalSurfaces'):
            try:
                repo = getattr(a, repo_name)
                if region_name in repo_keys(repo):
                    return True
            except Exception:
                pass

        for inst_name in repo_keys(a.instances):
            inst = a.instances[inst_name]
            for repo_name in ('sets', 'surfaces'):
                try:
                    repo = getattr(inst, repo_name)
                    if region_name in repo_keys(repo):
                        return True
                except Exception:
                    pass

        return False

    def capture_section_assignments(part):
        out = []
        try:
            for sa in part.sectionAssignments:
                section_name = str(sa.sectionName)
                region_name = None
                try:
                    rt = sa.region
                    if rt and len(rt) > 0:
                        region_name = str(rt[0])
                except Exception:
                    pass

                out.append({
                    'section_name': section_name,
                    'region_name': region_name
                })
        except Exception:
            pass
        return out

    def apply_sections(old_part_name, new_part, assignments):
        if not assignments:
            return

        unique_sections = []
        for x in assignments:
            if x['section_name'] not in unique_sections:
                unique_sections.append(x['section_name'])

        applied = 0

        # output stuff
        for x in assignments:
            rn = x.get('region_name')
            sn = x.get('section_name')
            if rn and rn in repo_keys(new_part.sets):
                try:
                    new_part.SectionAssignment(
                        region=new_part.sets[rn],
                        sectionName=sn
                    )
                    applied += 1
                except Exception as e:
                    warn(
                        'Section "%s" could not be assigned to recreated set "%s" on %s: %s'
                        % (sn, rn, old_part_name, e)
                    )

        # keep this
        # main part
        #
        # this part
        if applied == 0 and len(unique_sections) == 1:
            try:
                whole_name = '__AUTO_WHOLE_PART_SECTION_REGION__'
                if whole_name in repo_keys(new_part.sets):
                    del new_part.sets[whole_name]
                new_part.Set(name=whole_name, cells=new_part.cells[:])
                new_part.SectionAssignment(
                    region=new_part.sets[whole_name],
                    sectionName=unique_sections[0]
                )
                applied = 1
                log(
                    '  Section transfer: %s -> all cells of imported component'
                    % unique_sections[0]
                )
            except Exception as e:
                warn(
                    'Could not assign the single section from %s to the imported body: %s'
                    % (old_part_name, e)
                )

        if applied == 0 and len(unique_sections) > 1:
            warn(
                'Template part "%s" has multiple section assignments. '
                'The replacement STEP body has intentionally not been partitioned, '
                'so those multiple material regions cannot be transferred safely.'
                % old_part_name
            )

    #
    # calc part
    ####

    template_cae = os.path.abspath(template_cae)
    step_file = os.path.abspath(step_file)

    report_path = os.path.splitext(template_cae)[0] + '_STEP_TRANSFER_REPORT.txt'

    log('=' * 78)
    log('STEP -> ABAQUS TEMPLATE GEOMETRY TRANSFER')
    log('=' * 78)
    log('Template CAE : %s' % template_cae)
    log('STEP geometry: %s' % step_file)
    log('CAE updated in place: %s' % template_cae)
    log('Report       : %s' % report_path)
    log('Partitions   : NONE CREATED BY THIS SCRIPT')
    log('')

    if not os.path.exists(template_cae):
        raise Exception('Template CAE does not exist: %s' % template_cae)

    if not os.path.exists(step_file):
        raise Exception('STEP file does not exist: %s' % step_file)

    log('Opening template CAE...')
    openMdb(pathName=template_cae)

    # output stuff
    # keep this
    model_names = repo_keys(mdb.models)
    if not model_names:
        raise Exception('No models were found in the template CAE.')

    best_model_name = None
    best_model_score = -1

    for mn in model_names:
        mm = mdb.models[mn]
        score = 0
        try:
            score += len(mm.parts)
        except Exception:
            pass
        try:
            score += 2*len(mm.rootAssembly.instances)
        except Exception:
            pass
        try:
            score += 5*len(mm.boundaryConditions)
        except Exception:
            pass
        try:
            score += 5*len(mm.loads)
        except Exception:
            pass
        try:
            score += 3*len(mm.interactions)
        except Exception:
            pass

        if score > best_model_score:
            best_model_score = score
            best_model_name = mn

    model = mdb.models[best_model_name]
    assembly = model.rootAssembly

    log('Template model automatically selected: %s' % best_model_name)
    log('')

    old_instance_names = repo_keys(assembly.instances)
    if not old_instance_names:
        raise Exception('The selected template model has no assembly instances.')

    old_instance_bboxes = {}
    for name in old_instance_names:
        try:
            old_instance_bboxes[name] = instance_bbox(assembly.instances[name])
        except Exception as e:
            warn('Could not get template instance bounding box for %s: %s' % (name, e))

    if not old_instance_bboxes:
        raise Exception('Could not obtain any template instance geometry.')

    old_global_bb = _merge_bboxes(list(old_instance_bboxes.values()))

    # main part
    old_part_to_instances = {}
    for inst_name in old_instance_names:
        inst = assembly.instances[inst_name]
        try:
            part_name = str(inst.partName)
        except Exception:
            part_name = None

        if not part_name:
            # needed here
            for pn in repo_keys(model.parts):
                try:
                    if assembly.instances[inst_name].part == model.parts[pn]:
                        part_name = pn
                        break
                except Exception:
                    pass

        if part_name:
            old_part_to_instances.setdefault(part_name, []).append(inst_name)

    if not old_part_to_instances:
        raise Exception('Could not identify the Parts used by the template assembly.')

    old_part_names = sorted(old_part_to_instances.keys())

    log('Template geometric components found: %d' % len(old_part_names))
    for pn in old_part_names:
        log('  %s -> instance(s): %s' % (pn, ', '.join(old_part_to_instances[pn])))
    log('')

    # #
    #
    # calc part
    #
    # #

    attribute_regions = []
    required_region_names = set()

    def remember_region_name(value):
        name = None
        try:
            if isinstance(value, str):
                name = value
        except Exception:
            pass

        if not name:
            try:
                name = str(value.name)
            except Exception:
                pass

        if not name:
            try:
                if value and len(value) > 0:
                    first = value[0]
                    if isinstance(first, str):
                        name = str(first)
                    else:
                        try:
                            name = str(first.name)
                        except Exception:
                            pass
            except Exception:
                pass

        if name and name not in ('None', ''):
            required_region_names.add(name)
        return name

    # keep this
    try:
        for name in repo_keys(model.boundaryConditions):
            obj = model.boundaryConditions[name]
            rn = attr_region_name(obj)
            if rn:
                required_region_names.add(rn)
            attribute_regions.append({
                'category': 'BC',
                'name': name,
                'class': obj.__class__.__name__,
                'region': rn
            })
    except Exception:
        pass

    try:
        for name in repo_keys(model.loads):
            obj = model.loads[name]
            rn = attr_region_name(obj)
            if rn:
                required_region_names.add(rn)
            attribute_regions.append({
                'category': 'LOAD',
                'name': name,
                'class': obj.__class__.__name__,
                'region': rn
            })
    except Exception:
        pass

    # main part
    # needed here
    # this part
    region_attrs = (
        'region', 'surface', 'master', 'slave', 'main', 'secondary',
        'controlPoint', 'bodyRegion', 'tieRegion', 'region1', 'region2'
    )
    for repo_name in (
        'constraints', 'interactions', 'predefinedFields',
        'fieldOutputRequests', 'historyOutputRequests'
    ):
        try:
            repo = getattr(model, repo_name)
        except Exception:
            continue
        for obj_name in repo_keys(repo):
            try:
                obj = repo[obj_name]
            except Exception:
                continue
            for aname in region_attrs:
                try:
                    remember_region_name(getattr(obj, aname))
                except Exception:
                    pass

    log('Named regions directly referenced by analysis objects: %d' % len(required_region_names))
    if required_region_names:
        log('  ' + ', '.join(sorted(required_region_names)))
    log('')

    #
    # calc part
    ####

    old_part_capture = {}

    for pn in old_part_names:
        part = model.parts[pn]
        rep_inst_name = old_part_to_instances[pn][0]
        rep_inst = assembly.instances[rep_inst_name]
        rep_bb = old_instance_bboxes[rep_inst_name]

        # output stuff
        # keep this
        # main part
        p_sets = {}
        p_surfaces = {}
        section_assignments = capture_section_assignments(part)
        section_region_names = set(
            x.get('region_name') for x in section_assignments if x.get('region_name')
        )
        needed_part_internal = set(required_region_names) | section_region_names

        #
        # this part
        try:
            for name in repo_keys(rep_inst.sets):
                try:
                    p_sets[name] = capture_set(
                        rep_inst.sets[name],
                        {rep_inst_name: rep_bb},
                        rep_bb
                    )
                except Exception:
                    pass
        except Exception:
            pass

        try:
            for name in repo_keys(rep_inst.surfaces):
                try:
                    p_surfaces[name] = capture_surface(
                        rep_inst.surfaces[name],
                        {rep_inst_name: rep_bb},
                        rep_bb
                    )
                except Exception:
                    pass
        except Exception:
            pass

        #
        try:
            for name in repo_keys(part.allInternalSets):
                if name not in p_sets and name in needed_part_internal:
                    try:
                        # calc part
                        #
                        # output stuff
                        p_sets[name] = capture_set(
                            part.allInternalSets[name],
                            {},
                            rep_bb
                        )
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            for name in repo_keys(part.allInternalSurfaces):
                if name not in p_surfaces and name in needed_part_internal:
                    try:
                        p_surfaces[name] = capture_surface(
                            part.allInternalSurfaces[name],
                            {},
                            rep_bb
                        )
                    except Exception:
                        pass
        except Exception:
            pass

        old_part_capture[pn] = {
            'representative_instance': rep_inst_name,
            'bbox': rep_bb,
            'sets': p_sets,
            'surfaces': p_surfaces,
            'section_assignments': section_assignments
        }

    # keep this
    # main part
    root_sets = {}
    root_surfaces = {}

    try:
        for name in repo_keys(assembly.sets):
            try:
                root_sets[name] = capture_set(
                    assembly.sets[name], old_instance_bboxes, old_global_bb
                )
            except Exception:
                pass
    except Exception:
        pass

    try:
        for name in repo_keys(assembly.surfaces):
            try:
                root_surfaces[name] = capture_surface(
                    assembly.surfaces[name], old_instance_bboxes, old_global_bb
                )
            except Exception:
                pass
    except Exception:
        pass

    try:
        for name in repo_keys(assembly.allInternalSets):
            if name in required_region_names and name not in root_sets:
                try:
                    root_sets[name] = capture_set(
                        assembly.allInternalSets[name], old_instance_bboxes, old_global_bb
                    )
                except Exception:
                    pass
    except Exception:
        pass

    try:
        for name in repo_keys(assembly.allInternalSurfaces):
            if name in required_region_names and name not in root_surfaces:
                try:
                    root_surfaces[name] = capture_surface(
                        assembly.allInternalSurfaces[name], old_instance_bboxes, old_global_bb
                    )
                except Exception:
                    pass
    except Exception:
        pass

    log('Captured template assembly regions:')
    log('  Sets    : %d' % len(root_sets))
    log('  Surfaces: %d' % len(root_surfaces))
    log('')

    #
    log('Template boundary conditions / loads:')
    for x in attribute_regions:
        log(
            '  %-4s %-24s %-20s region=%s'
            % (x['category'], x['class'], x['name'], x['region'])
        )
    log('')

    # #
    #
    ##

    log('Opening STEP geometry...')
    geometry_file = mdb.openStep(fileName=step_file, scale=1.0)

    try:
        number_of_bodies = int(geometry_file.numberOfParts)
    except Exception:
        number_of_bodies = 0

    if number_of_bodies <= 0:
        raise Exception('Abaqus did not find any solid bodies in the STEP file.')

    log('STEP solid bodies reported by Abaqus: %d' % number_of_bodies)
    log('')

    imported_parts = []
    temp_instance_names = []

    for body_num in range(1, number_of_bodies + 1):
        part_name = '__AUTO_STEP_BODY_%03d__' % body_num

        #
        if part_name in repo_keys(model.parts):
            try:
                del model.parts[part_name]
            except Exception:
                pass

        log('Importing STEP body %d / %d...' % (body_num, number_of_bodies))

        p = model.PartFromGeometryFile(
            name=part_name,
            geometryFile=geometry_file,
            dimensionality=THREE_D,
            type=DEFORMABLE_BODY,
            bodyNum=body_num,
            combine=False,
            usePartNameFromFile=OFF
        )

        imported_parts.append(part_name)

        inst_name = '__AUTO_STEP_TEMP_%03d__' % body_num
        if inst_name in repo_keys(assembly.instances):
            del assembly.instances[inst_name]

        assembly.Instance(
            name=inst_name,
            part=p,
            dependent=ON
        )
        temp_instance_names.append(inst_name)

    try:
        assembly.regenerate()
    except Exception:
        pass

    new_body_bboxes = {}
    for part_name, inst_name in zip(imported_parts, temp_instance_names):
        new_body_bboxes[part_name] = instance_bbox(assembly.instances[inst_name])

    new_global_bb = _merge_bboxes(list(new_body_bboxes.values()))

    # #
    # keep this
    ##

    old_desc = {}
    for pn in old_part_names:
        old_desc[pn] = _part_match_descriptor(
            old_part_capture[pn]['bbox'],
            old_global_bb
        )

    new_desc = {}
    for part_name in imported_parts:
        new_desc[part_name] = _part_match_descriptor(
            new_body_bboxes[part_name],
            new_global_bb
        )

    #
    # this part
    #
    pairs = []
    for old_pn in old_part_names:
        for new_pn in imported_parts:
            pairs.append((
                _part_match_score(old_desc[old_pn], new_desc[new_pn]),
                old_pn,
                new_pn
            ))

    pairs.sort(key=lambda x: x[0])

    mapping = {}
    used_new = set()

    for score, old_pn, new_pn in pairs:
        if old_pn in mapping:
            continue
        if new_pn in used_new:
            continue

        mapping[old_pn] = {
            'new_part': new_pn,
            'score': score
        }
        used_new.add(new_pn)

        if len(mapping) >= min(len(old_part_names), len(imported_parts)):
            break

    log('')
    log('AUTOMATIC COMPONENT MATCHING')
    log('-' * 78)

    for old_pn in old_part_names:
        if old_pn in mapping:
            log(
                '  Template %-30s -> STEP %-24s  score=%.5f'
                % (
                    old_pn,
                    mapping[old_pn]['new_part'],
                    mapping[old_pn]['score']
                )
            )
            if mapping[old_pn]['score'] > 2.0:
                warn(
                    'Component match for "%s" has a relatively high score (%.5f). '
                    'Visually verify that body in the output CAE.'
                    % (old_pn, mapping[old_pn]['score'])
                )
        else:
            warn('No STEP body could be mapped to template component "%s".' % old_pn)

    extra_new = [x for x in imported_parts if x not in used_new]
    if extra_new:
        warn(
            'STEP contains %d extra body/bodies not matched to a template component: %s'
            % (len(extra_new), ', '.join(extra_new))
        )

    if len(imported_parts) != len(old_part_names):
        warn(
            'Template has %d geometric component Part(s), but STEP has %d solid body/bodies. '
            'The script will transfer what can be matched and leave extras clearly named.'
            % (len(old_part_names), len(imported_parts))
        )

    log('')

    ##
    #
    # output stuff
    #

    temp_by_part = dict(zip(imported_parts, temp_instance_names))

    for old_pn in old_part_names:
        if old_pn not in mapping:
            continue

        new_pn = mapping[old_pn]['new_part']
        new_part = model.parts[new_pn]
        temp_inst = assembly.instances[temp_by_part[new_pn]]

        old_bb = old_part_capture[old_pn]['bbox']
        new_bb = new_body_bboxes[new_pn]

        log('Rebuilding template regions on imported component for: %s' % old_pn)

        p_sets = old_part_capture[old_pn]['sets']
        for set_name in sorted(p_sets.keys()):
            ok, max_score, count = map_part_set_to_new_part(
                set_name,
                p_sets[set_name],
                old_bb,
                temp_inst,
                new_part,
                new_bb
            )

            if ok:
                log(
                    '  SET     %-32s mapped entities=%-4d maxScore=%.4f'
                    % (set_name, count, max_score)
                )
                if max_score > 2.5:
                    warn(
                        'Part set "%s" on "%s" mapped with a high geometry score %.4f.'
                        % (set_name, old_pn, max_score)
                    )

        p_surfaces = old_part_capture[old_pn]['surfaces']
        for surf_name in sorted(p_surfaces.keys()):
            ok, max_score, count = map_part_surface_to_new_part(
                surf_name,
                p_surfaces[surf_name],
                old_bb,
                temp_inst,
                new_part,
                new_bb
            )

            if ok:
                log(
                    '  SURFACE %-32s mapped faces=%-4d maxScore=%.4f'
                    % (surf_name, count, max_score)
                )
                if not p_surfaces[surf_name].get('sidedness_known'):
                    log(
                        '           note: Abaqus did not expose original surface sidedness; '
                        'fallback SIDE1 was used.'
                    )
                if max_score > 2.5:
                    warn(
                        'Part surface "%s" on "%s" mapped with a high geometry score %.4f.'
                        % (surf_name, old_pn, max_score)
                    )

        apply_sections(
            old_pn,
            new_part,
            old_part_capture[old_pn]['section_assignments']
        )

        log('')

    ##
    #
    # #

    for inst_name in temp_instance_names:
        try:
            if inst_name in repo_keys(assembly.instances):
                del assembly.instances[inst_name]
        except Exception:
            pass

    try:
        assembly.regenerate()
    except Exception:
        pass

    #
    # calc part
    #
    # output stuff
    #

    log('Replacing template instances with imported STEP geometry...')
    for old_pn in old_part_names:
        if old_pn not in mapping:
            continue

        new_pn = mapping[old_pn]['new_part']
        new_part = model.parts[new_pn]

        for inst_name in old_part_to_instances[old_pn]:
            try:
                inst = assembly.instances[inst_name]

                # main part
                try:
                    tr = inst.getTranslation()
                except Exception:
                    tr = (0.0, 0.0, 0.0)

                try:
                    rot = inst.getRotation()
                except Exception:
                    rot = None

                if _norm(tuple(tr)) > 1.0e-8:
                    warn(
                        'Template instance "%s" has a non-zero assembly translation %s. '
                        'The replacement keeps the template placement. Verify the STEP export '
                        'uses compatible component coordinates.'
                        % (inst_name, str(tr))
                    )

                if rot is not None:
                    try:
                        angle = float(rot[2])
                        if abs(angle) > 1.0e-8:
                            warn(
                                'Template instance "%s" has a non-zero assembly rotation. '
                                'The replacement keeps that placement; verify it visually.'
                                % inst_name
                            )
                    except Exception:
                        pass

                # needed here
                # this part
                inst.replace(
                    instanceOf=new_part,
                    applyConstraints=False
                )

                log('  Replaced instance: %s' % inst_name)

            except Exception as e:
                warn('Could not replace template instance "%s": %s' % (inst_name, e))

    try:
        assembly.regenerate()
    except Exception as e:
        warn('Assembly regeneration after instance replacement reported: %s' % e)

    #
    new_instance_bboxes = {}
    for inst_name in old_instance_names:
        if inst_name in repo_keys(assembly.instances):
            try:
                new_instance_bboxes[inst_name] = instance_bbox(assembly.instances[inst_name])
            except Exception:
                pass

    ##
    #
    #
    # keep this
    ##

    #
    # this part
    descriptor_cache.clear()

    log('')
    log('Rebuilding assembly-level named/picked regions...')
    log('-' * 78)

    for set_name in sorted(root_sets.keys()):
        ok, max_score, count = map_assembly_set(
            set_name,
            root_sets[set_name],
            assembly,
            old_instance_bboxes,
            new_instance_bboxes
        )
        if ok:
            log(
                '  SET     %-32s mapped entities=%-4d maxScore=%.4f'
                % (set_name, count, max_score)
            )
            if max_score > 2.5:
                warn(
                    'Assembly set "%s" mapped with a high geometry score %.4f.'
                    % (set_name, max_score)
                )

    for surf_name in sorted(root_surfaces.keys()):
        ok, max_score, count = map_assembly_surface(
            surf_name,
            root_surfaces[surf_name],
            assembly,
            old_instance_bboxes,
            new_instance_bboxes
        )
        if ok:
            log(
                '  SURFACE %-32s mapped faces=%-4d maxScore=%.4f'
                % (surf_name, count, max_score)
            )
            if not root_surfaces[surf_name].get('sidedness_known'):
                log(
                    '           note: original surface sidedness was not queryable; '
                    'fallback SIDE1 was used.'
                )
            if max_score > 2.5:
                warn(
                    'Assembly surface "%s" mapped with a high geometry score %.4f.'
                    % (surf_name, max_score)
                )

    #
    # calc part
    #
    # output stuff
    # keep this
    ##

    log('')
    log('Cleaning old template geometry Parts...')

    for old_pn in old_part_names:
        if old_pn not in mapping:
            continue

        new_pn = mapping[old_pn]['new_part']

        #
        try:
            if old_pn in repo_keys(model.parts):
                del model.parts[old_pn]
        except Exception as e:
            warn('Could not delete unused old template Part "%s": %s' % (old_pn, e))

        # this part
        try:
            if new_pn in repo_keys(model.parts) and old_pn not in repo_keys(model.parts):
                model.parts.changeKey(fromName=new_pn, toName=old_pn)
                log('  Imported Part renamed: %s -> %s' % (new_pn, old_pn))
        except Exception as e:
            warn(
                'Could not rename imported Part "%s" to "%s": %s'
                % (new_pn, old_pn, e)
            )

    #
    # calc part
    for idx, new_pn in enumerate(extra_new):
        if new_pn not in repo_keys(model.parts):
            continue
        try:
            extra_inst = 'STEP_EXTRA_BODY_%03d' % (idx + 1)
            if extra_inst not in repo_keys(assembly.instances):
                assembly.Instance(
                    name=extra_inst,
                    part=model.parts[new_pn],
                    dependent=ON
                )
                log('  Extra STEP body retained as instance: %s' % extra_inst)
        except Exception as e:
            warn('Could not instance extra STEP body "%s": %s' % (new_pn, e))

    try:
        assembly.regenerate()
    except Exception as e:
        warn('Final assembly regeneration reported: %s' % e)

    ####
    #
    #

    log('')
    log('BOUNDARY CONDITION / LOAD REGION CHECK')
    log('-' * 78)

    for x in attribute_regions:
        rn = x['region']
        if not rn:
            log(
                '  %-4s %-24s %-20s : no named region was exposed by Abaqus'
                % (x['category'], x['class'], x['name'])
            )
            continue

        exists = named_region_exists(model, rn)

        if exists:
            log(
                '  OK   %-4s %-24s %-20s region=%s'
                % (x['category'], x['class'], x['name'], rn)
            )
        else:
            warn(
                '%s "%s" (%s) references region "%s", but that named region '
                'could not be confirmed after geometry replacement.'
                % (x['category'], x['name'], x['class'], rn)
            )

    # main part
    # needed here
    log('')
    log('SYMMETRY BC SUMMARY')
    log('-' * 78)

    symmetry_found = False
    try:
        for name in repo_keys(model.boundaryConditions):
            obj = model.boundaryConditions[name]
            cls = obj.__class__.__name__
            upper = (name + ' ' + cls).upper()
            if (
                'XSYMM' in upper or
                'YSYMM' in upper or
                'ZSYMM' in upper or
                'SYMM' in upper
            ):
                symmetry_found = True
                log(
                    '  %s   type=%s   region=%s'
                    % (name, cls, attr_region_name(obj))
                )
    except Exception:
        pass

    if not symmetry_found:
        warn(
            'No symmetry BC could be identified by name/class in the selected template model. '
            'If the template uses symmetry, inspect the output CAE before analysis.'
        )

    # #
    #
    ##

    log('')
    log('Saving converted model...')
    mdb.save()

    log('')
    log('=' * 78)
    log('TRANSFER FINISHED')
    log('=' * 78)
    log('Updated CAE: %s' % template_cae)
    log('No partition operations were created by this script.')
    log('Warnings: %d' % len(warnings))
    log('')
    log('IMPORTANT FINAL CHECK:')
    log('Open the new CAE and visually display Loads + BCs. Confirm the pressure')
    log('surface and X/Y symmetry faces are on the intended new STEP geometry.')
    log('Face topology can legitimately change when CAD dimensions change, so this')
    log('visual check should be kept even after the automation is working reliably.')
    log('=' * 78)

    try:
        with open(report_path, 'w') as f:
            f.write('\n'.join(report_lines))
            f.write('\n')
    except Exception as e:
        print('Could not write report: %s' % e)

    return 0




def _abaqus_probe(template_cae, step_file):
    """Minimal diagnostic: open the CAE, open STEP, and import each STEP body into a temporary model. Never saves."""
    from abaqus import mdb, openMdb
    from abaqusConstants import THREE_D, DEFORMABLE_BODY, OFF

    def p(text):
        print(str(text))
        try:
            sys.stdout.flush()
        except Exception:
            pass

    template_cae = os.path.abspath(template_cae)
    step_file = os.path.abspath(step_file)

    p('=' * 72)
    p('MINIMAL ABAQUS 2023 CAE + STEP IMPORT PROBE')
    p('=' * 72)
    p('Python: %s' % sys.version.replace('\n', ' '))
    p('CAE : %s' % template_cae)
    p('STEP: %s' % step_file)

    p('STAGE 1/5 - importing Abaqus modules: OK')
    p('STAGE 2/5 - opening CAE...')
    openMdb(pathName=template_cae)
    p('STAGE 2/5 - opening CAE: OK')
    p('  Models: %s' % ', '.join([str(x) for x in mdb.models.keys()]))

    p('STAGE 3/5 - opening STEP...')
    geometry_file = mdb.openStep(fileName=step_file, scale=1.0)
    p('STAGE 3/5 - opening STEP: OK')
    n = int(geometry_file.numberOfParts)
    p('  STEP bodies reported: %d' % n)
    if n <= 0:
        raise Exception('STEP opened but Abaqus reported zero bodies.')

    probe_model_name = '__STEP_IMPORT_PROBE__'
    if probe_model_name in mdb.models.keys():
        del mdb.models[probe_model_name]
    probe_model = mdb.Model(name=probe_model_name)

    p('STAGE 4/5 - importing STEP bodies individually...')
    for body_num in range(1, n + 1):
        name = 'PROBE_BODY_%03d' % body_num
        p('  importing body %d/%d...' % (body_num, n))
        part = probe_model.PartFromGeometryFile(
            name=name,
            geometryFile=geometry_file,
            dimensionality=THREE_D,
            type=DEFORMABLE_BODY,
            bodyNum=body_num,
            combine=False,
            usePartNameFromFile=OFF
        )
        try:
            p('    OK: cells=%d faces=%d edges=%d' % (len(part.cells), len(part.faces), len(part.edges)))
        except Exception:
            p('    OK')

    p('STAGE 4/5 - STEP body imports: OK')
    p('STAGE 5/5 - probe complete. Nothing will be saved.')
    p('=' * 72)
    return 0


def _worker_entry():
    cae = os.environ.get(ENV_CAE, '')
    step = os.environ.get(ENV_STEP, '')
    live_log = os.environ.get(ENV_LOG, '')
    mode = os.environ.get(ENV_MODE, 'full').strip().lower()

    if not cae or not step:
        print('Missing worker environment variables.')
        return 2

    try:
        if mode == 'probe':
            return _abaqus_probe(cae, step)
        return _abaqus_worker(cae, step)
    except Exception as e:
        text = '\nFATAL ABAQUS ERROR\nMODE: %s\n%s\n\n%s\n' % (mode, str(e), traceback.format_exc())
        try:
            print(text)
            sys.stdout.flush()
        except Exception:
            pass
        if live_log:
            try:
                with open(live_log, 'a') as f:
                    f.write(text)
            except Exception:
                pass
        return 1


####
#
#
def main():
    if os.environ.get(WORKER_FLAG, '') == '1':
        return _worker_entry()
    return _normal_python_gui()


if __name__ == '__main__':
    sys.exit(main())
