# RDPWrapKit - The all-in-one installer for local RDP setup

### Quick Start
1. Download [RDPWrapKit-Setup.exe](https://github.com/cpdx4/RDPWrapKit/releases/latest/download/RDPWrapKit-Setup.exe)
2. Run it and choose **Typical Setup**
3. Open the new RDP shortcut on your desktop

  <img width="571" height="346" alt="image" src="https://github.com/user-attachments/assets/bf059039-6677-4105-be13-139f0a536f40" />

---
### Why this exists
- **Problem**: Local RDP setups are complex, prone to errors, and requires upkeep
- **Solution**: RDPWrapKit bundles the community's best tools into a single installer that _just works_

### What's in it
- ☑ [TermWrap](https://github.com/llccd/TermWrap) - Seamless updates (eliminating need for "rdpwrap.ini" files to be updated)
- ☑ Optional user creation during setup
- ☑ Proactively identify and auto-fix common RDP misconfigurations
- ☑ Creates ready‑to‑use RDP shortcuts on the desktop
- ☑ Fully transparent (open source) installer script (Inno Setup)
---
### How to install it:
1. Download [RDPWrapKit-Setup.exe](https://github.com/cpdx4/RDPWrapKit/releases/latest/download/RDPWrapKit-Setup.exe)
   - Your browser or Antivirus might need to be modified to allow it to download
   - It is virus free. If worried, you can inspect the [source code](https://github.com/cpdx4/RDPWrapKit/blob/main/RDPWrapKit.iss) or build it yourself (instructions at bottom)

2. Run the **RDPWrapKit-Setup.exe** app:

  - **For new installs**, use "Typical Setup" ("Install TermWrap" and "Create Users")
  - If you're already running TermWrap, choose "Use existing users" and skip installation
    <img width="598" height="464" alt="Installer_SetupOptions" src="https://github.com/user-attachments/assets/5f2e9b9b-2633-44d4-bdaf-2fdd40b8d74d" />
  
4. If you chose to "Create Users", you can create a new user/password (such as 'macro1')

      <img width="598" height="464" alt="Installer_CreateRDPUser" src="https://github.com/user-attachments/assets/d82f106e-7f08-4082-83e3-c344274a8d22" />
  

6. The install takes care of all the typical steps required to setup local RDP. Restart if prompted:

      <img width="571" height="346" alt="image" src="https://github.com/user-attachments/assets/bf059039-6677-4105-be13-139f0a536f40" />

8. Open the new RDP shortcut(s) on the desktop:

     <img width="571" height="346" alt="Installer_DesktopIcon" src="https://github.com/user-attachments/assets/15e7947b-9a11-42b3-b9c7-5f081ce2947f" />


---
### Support / Contact
- There is no implied warranty and unexpected results might occur
- Open an [issue](https://github.com/cpdx4/RDPWrapKit/issues) on this repository for questions, bugs, or enhancement requests

# Credits
- RDPWrapKit builds on the fantastic work of [TermWrap](https://github.com/llccd/TermWrap), which was inspired by [RDPWrapper](https://github.com/stascorp/rdpwrap/releases)
- Bee Swarm Simulator communities ([BSGH](https://discord.gg/bsgh) and [BSS Grinders](https://discord.gg/K5U3RdGXh6))


---
### To build this from source code
#### (Advanced. NOT required for most people):
1. Install [Inno Setup](https://github.com/jrsoftware/issrc/releases)
2. Download the [Source Code](https://github.com/cpdx4/RDPWrapKit/archive/refs/heads/main.zip)
3. Open `RDPWrapKit.iss` in the Inno Setup Compiler.
4. Compile the script (press F9 or use the Compile button).
5. The generated installer will appear in the `output/` folder.
