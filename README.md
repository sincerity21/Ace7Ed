# Ace7Ed

A fork of [GreenTrafficLight’s Ace7Ed](https://github.com/GreenTrafficLight/Ace7Ed) with new features.

### New features

- **Undo** text edits (**Main → Undo** or **Ctrl+Z**).
- **Export** to CSV with an optional **start** and **end** string number (**Options → Export…**).
- **Import** from CSV with an option to **overwrite existing strings** or only fill empty slots (**Options → Import…**).
- **Copy / paste to languages**: copy selected cells into one or more chosen target languages in one step.
- **Add an add-on** (**Options → Add an add-on**) to add plane/skin-related string entries (e.g. for mods).

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
