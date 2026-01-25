# RDPWrapKit

RDPWrapKit bundles RDPWrap and TermWrap payloads and an Inno Setup script to produce a Windows installer.

**Overview**
- Purpose: Build an easy-to-run installer that deploys RDPWrap and TermWrap components on Windows.
- Components: Inno Setup script and payload folders containing the binaries/resources.

**Repository Structure**
- RDPWrapKit.iss — Inno Setup script to build the installer. See [RDPWrapKit.iss](RDPWrapKit.iss)
- payload/ — Files included by the installer (RDPWrap, TermWrap). See [payload/](payload/)
- Output/ — Build artifacts produced by the Inno Setup compiler.

**Requirements**
- Windows (target platform for built installer)
- Inno Setup (recommended: Inno Setup 6) to compile the `.iss` script

**Build (create the installer)**
1. Install Inno Setup from the official source.
2. Open `RDPWrapKit.iss` in the Inno Setup Compiler.
3. Compile the script (press F9 or use the Compile button).
4. The generated installer will appear in the `Output/` folder.

**Install (end user)**
- Run the generated installer as Administrator and follow prompts.

**Updating payloads**
- Replace or update files inside `payload/RDPWrap/` and `payload/TermWrap/` as needed.
- Recompile `RDPWrapKit.iss` to produce a new installer containing the updated payload.

**Development / Contributing**
- Make changes, open issues for problems, or submit pull requests with improvements.
- When changing payload contents, document the source and version of any upstream binaries.

**License**
- Check for a `LICENSE` file in the repository. If none exists, please contact the maintainer or open an issue to clarify licensing.

**Support / Contact**
- Open an issue on this repository for questions, bugs, or enhancement requests.

---

Created for easy packaging of RDPWrap + TermWrap; customize `payload/` then recompile `RDPWrapKit.iss` to produce updated installers.
