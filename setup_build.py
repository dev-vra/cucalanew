import os
import platform
import subprocess

apps = ["consolidador.spec", "gerenciador.spec"]

for spec in apps:
    print(f"🔨 Building {spec} ...")
    cmd = ["pyinstaller", "--noconfirm", spec]
    subprocess.run(cmd, check=True)

print("✅ Build finished. Check 'dist/' folder.")