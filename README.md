# RDPWrapKit

RDPWrapKit bundles RDPWrap and TermWrap payloads and an Inno Setup script to produce a Windows installer.

**Overview**
- Purpose: Build an easy-to-run installer that deploys RDPWrap and TermWrap components on Windows.
- Components: Inno Setup script and payload folders containing the binaries/resources.

**How to Install**
- Run the installer and follow prompts.

### This app allows the install of RDP Wrapper with TermWrap and creating accounts
<img width="598" height="464" alt="Installer_SetupOptions" src="https://github.com/user-attachments/assets/5f2e9b9b-2633-44d4-bdaf-2fdd40b8d74d" />

### 1 or more Accounts can be created, or existing accounts can be linked
<img width="598" height="464" alt="Installer_CreateRDPUser" src="https://github.com/user-attachments/assets/d82f106e-7f08-4082-83e3-c344274a8d22" />

### The install takes care of all the typical steps required to setup local RDP
<img width="594" height="459" alt="Installer_Installing" src="https://github.com/user-attachments/assets/5e25e8d3-6ab2-4776-a47d-f42e9f079a2f" />

### A local RDP shortcut with embedded credentials is placed on the desktop for each user created
<img width="571" height="346" alt="Installer_DesktopIcon" src="https://github.com/user-attachments/assets/15e7947b-9a11-42b3-b9c7-5f081ce2947f" />


## If you want to build this yourself (Advanced, not required for most people):

**Requirements**
- Windows (target platform for built installer)
- Inno Setup (recommended: Inno Setup 6) to compile the `.iss` script

**Build (create the installer)**
1. Install Inno Setup from the official source.
2. Open `RDPWrapKit.iss` in the Inno Setup Compiler.
3. Compile the script (press F9 or use the Compile button).
4. The generated installer will appear in the `Output/` folder.

**Support / Contact**
- Open an issue on this repository for questions, bugs, or enhancement requests.

---

Created for easy packaging of RDPWrap + TermWrap
