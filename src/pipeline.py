import win32com.client
import os
import time
import subprocess
import shutil
import pymeshlab
import numpy as np

# ==========================================
# CONFIGURATION
# ==========================================
# Proje kök dizini (Kullanıcı değiştirebilir)
BASE_DIR = r"C:\3DPDF_Pipeline"

# Klasör Yapısı
DIRS = {
    "INPUT": os.path.join(BASE_DIR, "Input"),
    "TEMP": os.path.join(BASE_DIR, "Temp"),
    "OUTPUT": os.path.join(BASE_DIR, "Output")
}

# Kritik Dosyalar
FILES = {
    "JOB": os.path.join(DIRS["TEMP"], "job.txt"),
    "STL": os.path.join(DIRS["TEMP"], "temp_export.stl"),
    "LOG": os.path.join(DIRS["TEMP"], "worker_log.txt"),
    "PROCESSOR": os.path.join(DIRS["TEMP"], "Processor.ipt"),
    "TEX_TEMPLATE": os.path.join(DIRS["TEMP"], "render.tex"),
    "U3D_TARGET": os.path.join(DIRS["TEMP"], "model.u3d")
}

# ==========================================
# UTILS
# ==========================================
def init_directories():
    for d in DIRS.values():
        if not os.path.exists(d):
            os.makedirs(d)

def log(msg):
    try:
        print(f"[PIPELINE] {msg}")
    except:
        print(f"[PIPELINE] {msg.encode('ascii', 'replace').decode()}")

# ==========================================
# CORE LOGIC
# ==========================================
def execute_inventor_job(filename):
    """
    Inventor'ı 'Processor.ipt' dosyasını açarak tetikler.
    Bu yöntem COM Interface hatalarını önler.
    """
    full_path = os.path.join(DIRS["INPUT"], filename)
    log(f"Queueing Job: {filename}")
    
    # 1. Temizlik
    if os.path.exists(FILES["STL"]): os.remove(FILES["STL"])
    if os.path.exists(FILES["LOG"]): os.remove(FILES["LOG"])
    
    # 2. İş Emri Oluştur
    with open(FILES["JOB"], "w") as f:
        f.write(full_path)
    
    # 3. Inventor'ı Tetikle
    try:
        try:
            invApp = win32com.client.GetActiveObject("Inventor.Application")
        except:
            invApp = win32com.client.Dispatch("Inventor.Application")
            invApp.Visible = True
        
        invApp.SilentOperation = True
        
        # Processor dosyasını aç (Event Trigger çalışır)
        if not os.path.exists(FILES["PROCESSOR"]):
            log("CRITICAL: 'Processor.ipt' not found. Please create it in Temp folder.")
            return None

        oDoc = invApp.Documents.Open(FILES["PROCESSOR"])
        
        # 4. Sonuç Bekle (Smart Wait Loop)
        status = "TIMEOUT"
        for i in range(60): # 60 saniye zaman aşımı
            if os.path.exists(FILES["STL"]) and os.path.getsize(FILES["STL"]) > 100:
                status = "SUCCESS"
                break
            
            if os.path.exists(FILES["LOG"]):
                try:
                    with open(FILES["LOG"], 'r') as f:
                        content = f.read()
                        if "WARNING" in content:
                            status = "2D_SKIP"
                            break
                        if "ERROR" in content:
                            status = "ERROR"
                            log(f"   [WORKER LOG] {content.strip()}")
                            break
                except: pass
            time.sleep(1)
        
        oDoc.Close(True)
        invApp.SilentOperation = False
        
        if status == "SUCCESS":
            # STL'i güvenli bir isme taşı
            final_name = os.path.splitext(filename)[0] + ".stl"
            final_path = os.path.join(DIRS["TEMP"], final_name)
            if os.path.exists(final_path): os.remove(final_path)
            shutil.move(FILES["STL"], final_path)
            log(f"   [OK] STL Generated: {final_name}")
            return final_path
        elif status == "2D_SKIP":
            log("   [INFO] Skipped: File is 2D Drawing.")
            return None
        else:
            log(f"   [FAIL] Job finished with status: {status}")
            return None

    except Exception as e:
        log(f"   [SYSTEM ERROR] {e}")
        return None

def generate_pdf(stl_path, original_filename):
    """
    STL dosyasını MeshLab ile işler ve LaTeX ile PDF'e gömer.
    """
    file_name_no_ext = os.path.splitext(original_filename)[0]
    temp_u3d = stl_path.replace(".stl", ".u3d")
    
    log("   -> Mesh Processing & PDF Generation...")
    
    # MeshLab İşlemleri
    try:
        ms = pymeshlab.MeshSet()
        ms.load_new_mesh(stl_path)
        m = ms.current_mesh()
        
        # Ortalama ve Ölçekleme
        verts = m.vertex_matrix()
        center = verts.mean(axis=0)
        verts = verts - center
        max_dist = np.max(np.linalg.norm(verts, axis=1))
        if max_dist > 0:
            verts = verts * (10.0 / max_dist)
        
        ms_new = pymeshlab.MeshSet()
        ms_new.add_mesh(pymeshlab.Mesh(verts, m.face_matrix()))
        
        # Normalleri Düzelt
        try:
            ms_new.apply_filter('compute_vertex_normals')
            ms_new.apply_filter('orient_poly_faces_uniformly')
            ms_new.apply_filter('set_color_per_face', color1=pymeshlab.Color(200,200,200,255))
        except: pass
        
        ms_new.save_current_mesh(temp_u3d)
    except Exception as e:
        log(f"   [MESH ERROR] {e}")
        return

    # LaTeX Şablonu
    latex_content = r"""
\documentclass[a4paper]{article}
\usepackage{media9}
\usepackage[margin=1cm]{geometry}
\begin{document}
    \pagestyle{empty}
    \centerline{\Large \textbf{Project: """ + file_name_no_ext.replace("_", "\_") + r"""}}
    \vspace{1cm}
    \includemedia[
        width=0.9\linewidth, height=0.7\linewidth,
        activate=pageopen, 3Dmenu, 3Dtoolbar, 3Dlights=Day,
        3Dcoo=0 0 0, 3Droo=25, 3Dbg=1 1 1
    ]{}{model.u3d}
\end{document}
    """
    
    with open(FILES["TEX_TEMPLATE"], "w") as f: f.write(latex_content)
    
    # U3D'yi yerine taşı
    if os.path.exists(FILES["U3D_TARGET"]): os.remove(FILES["U3D_TARGET"])
    os.rename(temp_u3d, FILES["U3D_TARGET"])
    
    # PDF Derle
    cwd = os.getcwd()
    os.chdir(DIRS["TEMP"])
    subprocess.run(["pdflatex", "-interaction=nonstopmode", "render.tex"], capture_output=True)
    os.chdir(cwd)
    
    temp_pdf = os.path.join(DIRS["TEMP"], "render.pdf")
    if os.path.exists(temp_pdf):
        final_pdf = os.path.join(DIRS["OUTPUT"], file_name_no_ext + ".pdf")
        if os.path.exists(final_pdf): os.remove(final_pdf)
        os.rename(temp_pdf, final_pdf)
        log(f"   🎉 [SUCCESS] PDF Ready: {final_pdf}")
    else:
        log("   [FAIL] PDF compilation failed.")

def main():
    log("=== 3D PDF PIPELINE v1.0 ===")
    init_directories()
    
    files = [f for f in os.listdir(DIRS["INPUT"]) if f.endswith(('.rvt', '.dwg', '.dxf', '.ipt', '.iam'))]
    
    if not files:
        log("No supported files found in Input directory.")
    
    for f in files:
        stl_path = execute_inventor_job(f)
        if stl_path:
            generate_pdf(stl_path, f)
        log("-" * 30)

if __name__ == "__main__":
    main()