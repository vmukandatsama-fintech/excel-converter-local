"""
Build script for PLZ CV Engine
Creates a standalone executable for deployment
"""
import subprocess
import sys
import shutil
from pathlib import Path

def build_executable():
    """Build the standalone executable using PyInstaller"""
    print("=" * 70)
    print("PLZ CV ENGINE - BUILD SCRIPT")
    print("=" * 70)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print("✓ PyInstaller found")
    except ImportError:
        print("❌ PyInstaller not found")
        print("\nInstalling PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
        print("✓ PyInstaller installed")
    
    # Clean previous builds
    print("\nCleaning previous builds...")
    for folder in ['build', 'dist']:
        if Path(folder).exists():
            shutil.rmtree(folder)
            print(f"  ✓ Removed {folder}/")
    
    # Build the executable
    print("\nBuilding executable...")
    print("This may take a few minutes...\n")
    
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "PLZ_CV_Engine.spec", "--clean"],
        capture_output=True,
        text=True
    )
    
    if result.returncode != 0:
        print("❌ Build failed!")
        print("\nError output:")
        print(result.stderr)
        return False
    
    print("✓ Build completed successfully!")
    
    # Verify the executable exists
    exe_path = Path("dist/PLZ_CV_Engine.exe")
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"\n✓ Executable created: {exe_path}")
        print(f"  Size: {size_mb:.2f} MB")
        return True
    else:
        print("\n❌ Executable not found in dist/")
        return False

def create_deployment_package():
    """Create a deployment package with all necessary files"""
    print("\n" + "=" * 70)
    print("CREATING DEPLOYMENT PACKAGE")
    print("=" * 70)
    
    # Create deployment folder
    deploy_folder = Path("deployment_package")
    if deploy_folder.exists():
        shutil.rmtree(deploy_folder)
    deploy_folder.mkdir()
    
    # Copy executable
    exe_src = Path("dist/PLZ_CV_Engine.exe")
    exe_dst = deploy_folder / "PLZ_CV_Engine.exe"
    shutil.copy2(exe_src, exe_dst)
    print(f"✓ Copied: PLZ_CV_Engine.exe")
    
    # Copy config.ini template
    config_src = Path("config.ini")
    config_dst = deploy_folder / "config.ini"
    shutil.copy2(config_src, config_dst)
    print(f"✓ Copied: config.ini")
    
    # Create README for deployment
    readme_content = """PLZ CV ENGINE - DEPLOYMENT PACKAGE
=====================================

WHAT'S INCLUDED:
- PLZ_CV_Engine.exe : Main application
- config.ini : Configuration file
- DEPLOYMENT_INSTRUCTIONS.txt : Setup guide

REQUIREMENTS:
- Windows 10/11
- Microsoft Excel installed
- Network access to OneDrive sync folder
- (Optional) Email account for notifications

QUICK START:
1. Copy this entire folder to target PC
2. Edit config.ini to configure email settings
3. Double-click PLZ_CV_Engine.exe to start
4. Check system tray for application icon

For detailed instructions, see DEPLOYMENT_INSTRUCTIONS.txt
"""
    
    readme_path = deploy_folder / "README.txt"
    readme_path.write_text(readme_content, encoding='utf-8')
    print(f"✓ Created: README.txt")
    
    # Create detailed deployment instructions
    instructions_content = """PLZ CV ENGINE - DEPLOYMENT INSTRUCTIONS
=========================================

DEPLOYMENT STEPS:

1. PREREQUISITES (On Each Target PC)
   --------------------------------
   - Windows 10 or Windows 11
   - Microsoft Excel installed and activated
   - OneDrive sync configured
   - Network access to shared folder:
     Premium Leaf Zimbabwe\\v2SavannahTEST - Collection Vouchers

2. INSTALLATION (On Each Target PC)
   ---------------------------------
   Step 1: Copy deployment folder to desired location
           Example: C:\\Program Files\\PLZ CV Engine\\
   
   Step 2: Edit config.ini
           - For email notifications:
             * Set enabled = true
             * Configure SMTP server settings
             * Add recipient email addresses
           - For no email:
             * Set enabled = false
   
   Step 3: Test the configuration
           - Right-click PLZ_CV_Engine.exe
           - Select "Run as Administrator" (first time only)
           - Watch for system tray icon
           - Check activity_log.txt for status

3. EMAIL CONFIGURATION (Optional)
   ------------------------------
   For Gmail:
   1. Enable 2-factor authentication on sender account
   2. Generate App Password at: https://myaccount.google.com/apppasswords
   3. In config.ini:
      smtp_server = smtp.gmail.com
      smtp_port = 587
      sender_email = your.email@gmail.com
      sender_password = (paste app password here)
      recipients = recipient1@example.com, recipient2@example.com

   For Outlook/Office 365:
      smtp_server = smtp.office365.com
      smtp_port = 587
      sender_email = your.email@outlook.com
      sender_password = (your password or app password)

4. AUTO-START ON BOOT (Optional)
   -----------------------------
   To make the application start automatically:
   
   Method 1 - Task Scheduler (Recommended):
   1. Open Task Scheduler
   2. Create Basic Task
   3. Name: PLZ CV Engine
   4. Trigger: At log on
   5. Action: Start a program
   6. Program: (path to PLZ_CV_Engine.exe)
   7. Check "Run with highest privileges"
   
   Method 2 - Startup Folder:
   1. Press Win+R, type: shell:startup
   2. Create shortcut to PLZ_CV_Engine.exe
   3. Place shortcut in opened folder

5. FOLDER STRUCTURE
   ----------------
   The application automatically creates and monitors:
   
   C:\\Users\\[Username]\\Premium Leaf Zimbabwe\\v2SavannahTEST - Collection Vouchers\\
   ├── Json\\                    # Input folder (monitored)
   ├── Templates\\               # Excel template location
   ├── Populated Template\\      # Working folder (temporary files)
   └── PDF CVs\\                 # Output folder (generated PDFs)

6. TESTING
   -------
   1. Place a JSON file in the Json folder
   2. Watch for PDF generation in PDF CVs folder
   3. Check email inbox (if configured)
   4. Review activity_log.txt for processing details

7. MONITORING & LOGS
   -----------------
   - System tray icon shows application status
   - Right-click icon to check status or open PDF folder
   - activity_log.txt records all processing activities
   - Logs are created in the same folder as the .exe

8. TROUBLESHOOTING
   ---------------
   Problem: Application won't start
   Solution: Run as Administrator, check Windows Defender

   Problem: PDF not generated
   Solution: Verify Excel is installed, check template exists

   Problem: Email not sending
   Solution: Test config with test_email.py, verify credentials

   Problem: JSON files not detected
   Solution: Check OneDrive sync status, verify folder permissions

9. UNINSTALLATION
   --------------
   1. Right-click system tray icon → Exit
   2. Remove Task Scheduler task (if created)
   3. Delete application folder
   4. OneDrive folders remain - delete manually if needed

10. SUPPORT
    -------
    - Check activity_log.txt for error details
    - Review console output (if running in console mode)
    - Contact IT support with log files

SECURITY NOTES:
- config.ini contains email password - protect this file
- Use app passwords, never main account passwords
- Limit recipients to authorized personnel only
- Regularly review activity logs for anomalies
"""
    
    instructions_path = deploy_folder / "DEPLOYMENT_INSTRUCTIONS.txt"
    instructions_path.write_text(instructions_content, encoding='utf-8')
    print(f"✓ Created: DEPLOYMENT_INSTRUCTIONS.txt")
    
    # Create a batch file for easy startup
    batch_content = """@echo off
echo Starting PLZ CV Engine...
start "" "%~dp0PLZ_CV_Engine.exe"
"""
    batch_path = deploy_folder / "Start_PLZ_CV_Engine.bat"
    batch_path.write_text(batch_content)
    print(f"✓ Created: Start_PLZ_CV_Engine.bat")
    
    print(f"\n✓ Deployment package created: {deploy_folder.absolute()}")
    print(f"\nPackage contents:")
    for item in deploy_folder.iterdir():
        size = item.stat().st_size
        if size > 1024 * 1024:
            size_str = f"{size / (1024 * 1024):.2f} MB"
        elif size > 1024:
            size_str = f"{size / 1024:.2f} KB"
        else:
            size_str = f"{size} bytes"
        print(f"  - {item.name:40s} {size_str:>12s}")
    
    return deploy_folder

def main():
    """Main build process"""
    try:
        # Build the executable
        if not build_executable():
            print("\n❌ Build failed. Deployment package not created.")
            return
        
        # Create deployment package
        deploy_folder = create_deployment_package()
        
        print("\n" + "=" * 70)
        print("BUILD COMPLETE!")
        print("=" * 70)
        print(f"\nDeployment package ready at:")
        print(f"  {deploy_folder.absolute()}")
        print(f"\nTo deploy to other PCs:")
        print(f"  1. Copy the entire '{deploy_folder.name}' folder")
        print(f"  2. Follow instructions in DEPLOYMENT_INSTRUCTIONS.txt")
        print(f"  3. Configure config.ini on each PC")
        print(f"  4. Run PLZ_CV_Engine.exe")
        print("\n" + "=" * 70)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
