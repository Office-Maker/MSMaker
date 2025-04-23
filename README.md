# ⚠️ Windows Defender False Positives – Explanation

## Why Does Windows Defender Flag This Tool?

Windows Defender (and other antivirus software) may detect this tool as:

- `PUA:Win32/Puwaders.C!ml`
- `Trojan:Script/Wacatac.B!ml`

These are **false positives** caused by:

- Windows activation components (e.g., modifying `sppsvc`, `ClipSVC`, or registry keys)
- Batch files (`*.bat` / `*.cmd`) being inherently powerful (and thus often flagged)

> 🔹 **This does NOT mean the file is malicious.**  
> 🔹 The code is open-source and safe if downloaded from the official repository.

---

## How to Resolve This?

### ✅ Option 1: Use the "Lite" Version (Recommended)

If you only need Microsoft Office activation, download:  
**[MSMaker-LITE](https://github.com/Office-Maker/MSMaker-LITE)**

- Excludes Windows activation (no false positives in windows security)

---

### 🔧 Option 2: Add a Defender Exclusion

1. Open **Windows Security** → **Virus & Threat Protection** → **Manage Settings**  
2. Under **Exclusions**, add the folder containing `MSMaker.bat`
---

### ⏳ Option 3: Temporarily Disable Real-Time Protection (not recommended)

1. Turn off **Real-time protection** (Windows Security → Virus & Threat Protection)  
2. Run the script  
3. Re-enable protection afterward

---

## Note!

- No matter which download you choose, and even if Windows Security is disabled when running the batch file,  
  Windows will display a warning like **"Windows protected your PC"**.  
  This **always** occurs when running batch files downloaded from the internet and has nothing to do with potential risks or malware detections.  
  Since batch scripts are very powerful, Windows will **always** warn you before running them.

> 🔹 **This is normal. Proceed if you trust the source.**

---

## 📜 Legal Disclaimer & Agreement

By downloading or using this software, you agree that:

- This is for **educational/testing purposes only**
- You take **full responsibility** for any legal or technical consequences
- **Microsoft’s Terms of Service prohibit unauthorized activation** — use at your own risk
- The developers are **not liable** for misuse, damages, or compliance violations

> ⚠️ **Warning:** Distributing or using this tool in commercial environments may **violate Microsoft’s licensing terms**.  
> 💡 Purchase genuine licenses for legal use.
