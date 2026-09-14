import os
import sys
import zipfile
import subprocess
import shutil

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WWW_DIR = os.path.join(ROOT_DIR, 'android', 'app', 'src', 'main', 'assets', 'www')
BASE_APK = os.path.join(ROOT_DIR, 'BaoCaoDoanhSo_TeamCamGiang.apk')
TEMP_UNSIGNED = os.path.join(ROOT_DIR, 'temp_unsigned.apk')
OUT_DIR = os.path.join(ROOT_DIR, 'out_apk')
TOOLS_DIR = os.path.join(ROOT_DIR, 'tools')
SIGNER_JAR = os.path.join(TOOLS_DIR, 'uber-apk-signer.jar')
JAVA_EXE = os.path.join(TOOLS_DIR, 'jre', 'bin', 'java.exe')

def main():
    print("=== DANG DONG GOI VA CAP NHAT APK (PORTABLE SIGNER) ===")
    if not os.path.exists(BASE_APK):
        print(f"Error: Base APK not found at {BASE_APK}")
        sys.exit(1)
        
    if not os.path.exists(JAVA_EXE) or not os.path.exists(SIGNER_JAR):
        print(f"Error: Java or Signer tool not found in {TOOLS_DIR}")
        sys.exit(1)

    print("1. Dong goi ma nguon web moi nhat vao APK...")
    with zipfile.ZipFile(BASE_APK, 'r') as zin, zipfile.ZipFile(TEMP_UNSIGNED, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename.startswith('META-INF/'):
                continue
            if item.filename.startswith('assets/www/'):
                rel_name = item.filename[len('assets/www/'):]
                local_path = os.path.join(WWW_DIR, rel_name)
                if os.path.exists(local_path):
                    with open(local_path, 'rb') as f:
                        zout.writestr(item, f.read())
                    continue
            zout.writestr(item, zin.read(item.filename))

    print("2. Ky so (ZipAlign & APK Signature Scheme v2/v3)...")
    if os.path.exists(OUT_DIR):
        shutil.rmtree(OUT_DIR)
    os.makedirs(OUT_DIR, exist_ok=True)

    cmd = [
        JAVA_EXE, '-jar', SIGNER_JAR,
        '-a', TEMP_UNSIGNED,
        '-o', OUT_DIR,
        '--allowResign'
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print("Signer Error:\n", res.stderr or res.stdout)
        sys.exit(1)

    signed_name = 'temp_unsigned-aligned-debugSigned.apk'
    signed_path = os.path.join(OUT_DIR, signed_name)
    if not os.path.exists(signed_path):
        print("Error: Signed APK not generated.")
        sys.exit(1)

    print("3. Cap nhat cac file APK dau ra...")
    apk1 = os.path.join(ROOT_DIR, 'BaoCaoDoanhSo_TeamCamGiang.apk')
    apk2 = os.path.join(ROOT_DIR, 'BaoCaoThucDat.apk')
    shutil.copy2(signed_path, apk1)
    shutil.copy2(signed_path, apk2)

    # Clean up
    if os.path.exists(TEMP_UNSIGNED):
        os.remove(TEMP_UNSIGNED)
    if os.path.exists(OUT_DIR):
        shutil.rmtree(OUT_DIR)

    print("========================================================")
    print(" DA CAP NHAT THANH CONG 2 FILE APK MOI:")
    print(" -> BaoCaoDoanhSo_TeamCamGiang.apk")
    print(" -> BaoCaoThucDat.apk")
    print("========================================================")

if __name__ == '__main__':
    main()
