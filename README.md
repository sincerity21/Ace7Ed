# Ace7Ed

A fork of [GreenTrafficLight’s Ace7Ed](https://github.com/GreenTrafficLight/Ace7Ed) with new features.

### New features

- **Open single or selected languages** (**Main → Open single language…**): load only **Cmn.dat** plus one or more language DATs instead of all 13. Pick Cmn.dat first, then choose language file(s)—use **Ctrl+click** to select multiple (e.g. B and M). Useful when editing only a subset of languages. **Main → Open Folder** still loads the full set (Cmn + all 13 DATs).
- **Undo** text edits (**Main → Undo** or **Ctrl+Z**).
- **Export** to CSV with an optional **start** and **end** string number (**Options → Export…**).
- **Import** from CSV with an option to **overwrite existing strings** or only fill empty slots (**Options → Import…**).
- **Add a new plane** (**Options → Add a new plane**) to add plane/skin-related string entries (e.g. for mods).
- **Search** strings in the editor: use the **Search** box on the main form and choose a mode (**Number**, **ID**, or **Text**). Number matches the string index exactly; ID and Text filter by keyword (case-insensitive). Press **Enter** in the search box to apply the filter; clear the box to show all strings again.

### Building and running

- **Requirements:** .NET 8.0 (Windows).
- Open **Ace7Ed.sln** in Visual Studio (or use the .NET CLI), build the solution, and run the **Ace7Ed** project.

### Credits & third‑party code

Third‑party components and their licenses are described in **[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)**.

- **CUE4Parse**
  - Unreal Engine archive / package parsing library by FabianFG and contributors.
  - Repository: [FabianFG/CUE4Parse](https://github.com/FabianFG/CUE4Parse/)
  - Licensed under the [Apache License 2.0](https://github.com/FabianFG/CUE4Parse/blob/master/LICENSE).
  - This project uses CUE4Parse (directly or via bundled tools) in accordance with Apache‑2.0 and retains the original LICENSE and NOTICE files.

- **Ace7Ed and Ace7‑Localization‑Format (GreenTrafficLight)**
  - Original C# tooling and localization format for ACE COMBAT 7 by GreenTrafficLight.
  - Repositories:
    - [GreenTrafficLight/Ace7Ed](https://github.com/GreenTrafficLight/Ace7Ed)
    - [GreenTrafficLight/Ace7-Localization-Format](https://github.com/GreenTrafficLight/Ace7-Localization-Format/)
  - This fork builds on these projects and adds the features listed above.

- **This editor**
  - Enhancements and new features in this fork were vibe‑coded with Cursor Pro.
