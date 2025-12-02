# Open Source 3D PDF Conversion Pipeline for Autodesk Inventor

An automated pipeline designed to convert various CAD formats (`.ipt`, `.iam`, `.rvt`, `.dwg`) into interactive **3D PDFs** using Autodesk Inventor API, Python, and LaTeX. This project serves as an open-source alternative to commercial tools like Tetra4D.

## 🚀 Features

- **Multi-Format Support:** Handles Native Inventor files, Revit BIM models, and AutoCAD 3D DWGs.
- **Fail-Safe Architecture:** Uses a "Job Ticket" system to decouple Python from Inventor, preventing COM Interface errors.
- **Smart 2D Detection:** Automatically skips 2D DWG layouts to prevent pipeline clogging.
- **AnyCAD Integration:** Leverages Inventor's AnyCAD technology for high-fidelity Revit imports without data loss.
- **Automated Mesh Processing:** Integrates MeshLab for geometry normalization and optimization.

## 🛠️ Prerequisites

1.  **Autodesk Inventor Professional** (2024 or later recommended).
2.  **Python 3.10+**
3.  **MeshLab** (Ensure `pymeshlab` is compatible).
4.  **MiKTeX** or **TeX Live** (For `pdflatex` compilation).

## 📦 Installation

1.  Clone the repository:
    ```bash
    git clone [https://github.com/AhmetBerkeKaya/inventor-3dpdf-pipeline.git](https://github.com/yourusername/inventor-3dpdf-pipeline.git)
    ```
2.  Install Python dependencies:
    ```bash
    pip install -r requirements.txt
    ```
3.  **Critical Setup:**
    * Open Autodesk Inventor.
    * Create a new Part file named `Processor.ipt`.
    * Add an iLogic Rule named `WorkerBot` and paste the content from `resources/WorkerBot.vb`.
    * Set an **Event Trigger** for this rule on "After Open Document".
    * Save `Processor.ipt` into your `Temp` directory (defined in config).

## 🏃 Usage

Run the pipeline script:

```bash
python src/pipeline.py