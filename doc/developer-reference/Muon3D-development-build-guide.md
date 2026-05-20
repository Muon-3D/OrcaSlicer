# Muon3D OrcaSlicer Development Build Guide

This guide is for day-to-day Muon3D development on the Windows Visual Studio build of OrcaSlicer. It is intentionally narrower than the upstream build guide.

## Assumptions

- Repo path:

  ```bat
  D:\Dropbox\Muon3D_SharedFolder\Slicing\OrcaSlicer\OrcaSlicer
  ```

- Existing CMake build directory:

  ```bat
  build
  ```

- Generator:

  ```text
  Visual Studio 17 2022
  ```

- Build configuration:

  ```text
  Release
  ```

Run the commands below from a Visual Studio Developer Command Prompt or a normal command prompt where CMake and MSBuild are available.

## Important Targets

On Windows, the target names are easy to misunderstand:

```text
OrcaSlicer          -> main application DLL target
OrcaSlicer_app_gui  -> Windows GUI launcher executable target
INSTALL             -> copies built files into the staged app folder
```

`OrcaSlicer` builds:

```text
build\src\Release\OrcaSlicer.dll
```

`OrcaSlicer_app_gui` builds:

```text
build\src\Release\orca-slicer.exe
```

The staged folder is:

```text
build\OrcaSlicer\
```

This folder is not automatically updated just because the compile target was built. Run `INSTALL` when you need the staged copy refreshed.

## Quick Development Build

Use this when changing normal application, slicer, config, model, or GUI code and testing directly from the build output folder:

```bat
cd /d D:\Dropbox\Muon3D_SharedFolder\Slicing\OrcaSlicer\OrcaSlicer
cmake --build build --config Release --target OrcaSlicer
```

Then run:

```bat
build\src\Release\orca-slicer.exe
```

Even if `orca-slicer.exe` did not get a new timestamp, the application code may still have updated through:

```text
build\src\Release\OrcaSlicer.dll
```

## Full GUI Build

Use this when you want the normal GUI launcher target built as well:

```bat
cd /d D:\Dropbox\Muon3D_SharedFolder\Slicing\OrcaSlicer\OrcaSlicer
cmake --build build --config Release --target OrcaSlicer_app_gui
```

`OrcaSlicer_app_gui` depends on `OrcaSlicer`, so it should build the main DLL first if it is out of date.

Run:

```bat
build\src\Release\orca-slicer.exe
```

## Full Build Plus Staging

Use this when you want to test from the staged folder or share the staged build:

```bat
cd /d D:\Dropbox\Muon3D_SharedFolder\Slicing\OrcaSlicer\OrcaSlicer
cmake --build build --config Release --target OrcaSlicer_app_gui
cmake --build build --config Release --target INSTALL
```

Then run:

```bat
build\OrcaSlicer\orca-slicer.exe
```

After staging, check these files if you need to confirm what updated:

```text
build\src\Release\OrcaSlicer.dll
build\src\Release\orca-slicer.exe
build\OrcaSlicer\OrcaSlicer.dll
build\OrcaSlicer\orca-slicer.exe
```

For most code changes, `OrcaSlicer.dll` is the important timestamp.

## Incremental Builds

`cmake --build` is incremental. It should rebuild only the changed source files and the targets affected by those changes.

Common behavior:

```text
Changed .cpp file        -> recompiles that file and relinks affected target
Changed header file      -> recompiles files that include it, then relinks
Changed CMakeLists.txt   -> CMake regenerates, then builds affected targets
Changed resources        -> may require INSTALL to update staged resources
```

The slow part is often linking `OrcaSlicer.dll`, not recompiling the whole codebase.

## Which Command Should I Use?

For fast iteration:

```bat
cmake --build build --config Release --target OrcaSlicer
```

For normal GUI confidence:

```bat
cmake --build build --config Release --target OrcaSlicer_app_gui
```

For testing from `build\OrcaSlicer`:

```bat
cmake --build build --config Release --target OrcaSlicer_app_gui
cmake --build build --config Release --target INSTALL
```

For a clean rebuild of the app target, if incremental state seems wrong:

```bat
cmake --build build --config Release --target OrcaSlicer --clean-first
```

Use clean rebuilds sparingly because they are much slower.

## Notes for Excluded Area Development

The 3D excluded-area work touches both core and GUI-side code, including files such as:

```text
src\libslic3r\Model.cpp
src\libslic3r\PrintConfig.cpp
src\slic3r\GUI\PartPlate.cpp
src\slic3r\GUI\Tab.cpp
src\slic3r\GUI\Plater.cpp
src\slic3r\GUI\GUI.cpp
```

For that kind of change, the recommended development loop is:

```bat
cmake --build build --config Release --target OrcaSlicer_app_gui
build\src\Release\orca-slicer.exe
```

When ready to test the staged folder:

```bat
cmake --build build --config Release --target INSTALL
build\OrcaSlicer\orca-slicer.exe
```

## Troubleshooting

If `build\OrcaSlicer\orca-slicer.exe` did not update, run:

```bat
cmake --build build --config Release --target INSTALL
```

If `orca-slicer.exe` did not update but your code changed, check:

```text
build\src\Release\OrcaSlicer.dll
```

If the wrong executable is being run, prefer launching one of these explicitly:

```bat
build\src\Release\orca-slicer.exe
build\OrcaSlicer\orca-slicer.exe
```

If the build appears stale after CMake changes, build once with:

```bat
cmake --build build --config Release --target OrcaSlicer_app_gui --clean-first
```

