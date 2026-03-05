# PLZ CV ENGINE - DEPLOYMENT GUIDE

## ✅ Deployment Package

The current deployment package is in:
**`deployment_package/`**

---

## 📦 Package Contents

```
deployment_package/
├── PLZ_CV_Engine.exe              # Standalone executable
├── config.ini                     # Configuration file
├── Start_PLZ_CV_Engine.bat        # Quick start script
├── README.txt                     # Quick reference
└── DEPLOYMENT_INSTRUCTIONS.txt    # Detailed guide
```

---

## 🚀 Quick Deployment Steps

### **Step 1: Distribute Package**
1. Copy the **entire `deployment_package/` folder** to each target PC
2. Recommended location: `C:\Program Files\PLZ CV Engine\`
3. Or use a shared network drive for easy distribution

### **Step 2: Configure Email (Each PC)**
Edit `config.ini`:
```ini
[EMAIL]
enabled = true
smtp_server = smtp.gmail.com
smtp_port = 587
sender_email = your-email@example.com
sender_password = your-app-password
recipients = vmukandatsama@gmail.com, vmukandatsama@premiumzimbabwe.com
```

**💡 Tip:** You can use the SAME email credentials on all PCs (already configured)

### **Step 3: Launch Application**
- Double-click **`PLZ_CV_Engine.exe`** OR
- Double-click **`Start_PLZ_CV_Engine.bat`**
- Application runs in system tray (look for icon)

### **Step 4: Verify**
✓ Check system tray for PLZ icon
✓ Verify folder created: `C:\Users\[Username]\Premium Leaf Zimbabwe\...`
✓ Test by placing JSON file in `Json/` folder

---

## 🔄 Auto-Start Setup (Optional)

### Method A: Task Scheduler (Recommended)
```
1. Win+R → taskschd.msc
2. Create Basic Task → Name: "PLZ CV Engine"
3. Trigger: At log on
4. Action: Start program → Browse to PLZ_CV_Engine.exe
5. ✓ Run with highest privileges
```

### Method B: Startup Folder (Simple)
```
1. Win+R → shell:startup
2. Create shortcut to PLZ_CV_Engine.exe
3. Paste shortcut in opened folder
```

---

## 🌐 Multi-PC Deployment Strategy

### **Option 1: Manual Copy**
- Copy package to USB drive
- Install on each PC individually
- Best for: < 10 PCs

### **Option 2: Network Share**
```
1. Place package on shared drive (\\FILES\Software\PLZ_CV_Engine)
2. Users copy to local drive and run
3. Best for: 10-50 PCs
```

### **Option 3: Group Policy (Enterprise)**
```
1. Create MSI installer (advanced)
2. Deploy via Group Policy
3. Best for: 50+ PCs
```

---

## ⚙️ Configuration Management

### **Same Config for All PCs** (Recommended)
- All PCs use the same email account
- All PCs send to the same recipients
- Less maintenance, centralized monitoring

### **Different Config Per PC**
Each PC can have unique settings in config.ini:
- Different sender emails
- Different recipients
- Enabled/disabled independently

---

## 📝 Testing Checklist

On each deployed PC, verify:
- [ ] Application starts without errors
- [ ] System tray icon visible
- [ ] Folders created automatically
- [ ] JSON file detection works
- [ ] PDF generation successful
- [ ] Email sending works (if enabled)
- [ ] Activity log created

---

## 🔒 Security Considerations

1. **Email Credentials**
   - Use App Passwords (not main password)
   - Consider dedicated service account
   - Protect config.ini file

2. **File Permissions**
   - Ensure OneDrive sync permissions
   - Template file must be accessible
   - Output folder needs write access

3. **Firewall**
   - Allow outbound SMTP (port 587)
   - No inbound ports needed

---

## 🛠️ Maintenance

### **Update Deployment**
1. Rebuild: `python build_deployment.py`
2. Redistribute new package
3. Stop old version, replace .exe
4. Keep existing config.ini

### **View Logs**
- Check `activity_log.txt` next to .exe
- Grow over time - can be deleted safely
- Monitor for errors

### **Email Config Changes**
- Edit config.ini on each PC OR
- Deploy updated config.ini centrally

---

## 📞 Support

### **Common Issues**

**Application won't start:**
- Run as Administrator
- Check Windows Defender exclusions
- Verify Excel is installed

**PDF not generated:**
- Check template exists in OneDrive
- Verify Excel installation
- Review activity_log.txt

**Email not sending:**
- Verify credentials in config.ini
- Check firewall/antivirus blocking SMTP

**Files not detected:**
- Verify OneDrive sync status
- Check folder permissions
- Look for "modified" events in logs

---

## 📊 Deployment Summary

| Component | Details |
|-----------|---------|
| **Executable Size** | Depends on build |
| **Dependencies** | None (all bundled) |
| **Excel Required** | Yes |
| **Internet Required** | Only for email |
| **Admin Rights** | First run only |
| **OS Support** | Windows 10/11 |

---

## 🎯 Next Steps

1. ✅ Package created
2. ⏩ Test on one PC first
3. ⏩ Verify all functionality
4. ⏩ Deploy to remaining PCs
5. ⏩ Set up auto-start (optional)
6. ⏩ Monitor logs for first week

---

**Deployment Package Location:**
`C:\Users\VictorMukandatsama\Development\ExcelConverter\deployment_package\`

Ready to deploy! 🚀
