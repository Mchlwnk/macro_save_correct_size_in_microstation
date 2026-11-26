# Installation

This page describes how to install the macro in MicroStation V8i. The same
steps work for V8i SELECTseries 1–4.

## 1. Get the source file

Download `src/PrintCorrectSize.bas` from the repository. Keep the file
somewhere easy to find, for example your Desktop.

The `.bas` file is a plain VBA module. It must be imported into a VBA project
inside MicroStation. MicroStation stores VBA projects in `.mvba` files.

## 2. Find a folder for macros

MicroStation looks for VBA projects in the folders listed by the configuration
variable `MS_VBASEARCHDIRECTORIES`. A common location is:

```text
...\Workspace\Standards\Macros\vba
```

or, in a default V8i installation:

```text
C:\ProgramData\Bentley\MicroStation V8i (SELECTseries)\WorkSpace\Standards\Macros\vba
```

If the folder does not exist, you can create it. You can also use any folder
and add it to `MS_VBASEARCHDIRECTORIES`:

1. In MicroStation, open **Workspace > Configuration**.
2. Find `MS_VBASEARCHDIRECTORIES`.
3. Add your folder to the list.

For a first test it is enough to save the project anywhere and load it
manually with the Project Manager.

## 3. Create the VBA project

1. In MicroStation, open **Utilities > Macro > Project Manager**.
   In some V8i versions the menu is **Tools > VBA > Project Manager**.
2. Click **New Project** (the new-project icon).
3. Name the project:

   ```text
   PrintCorrectSize
   ```

4. Save it as `PrintCorrectSize.mvba` in the macro folder from step 2.
5. The project appears in the Project Manager list.

Alternative: open **Utilities > Key-in** and enter

```text
vba create PrintCorrectSize.mvba
```

This creates and loads the empty project. Then continue with the import step
below.

The project name matters: the example key-in uses
`[PrintCorrectSize]`. If you choose another name, use that name in the key-in.

## 4. Import the module

1. Select the new project in the Project Manager.
2. Click **Edit** (or **Show VBA Editor**) to open the Microsoft Visual Basic
   editor.
3. In the VBA editor, choose **File > Import File...**.
4. Select `PrintCorrectSize.bas`.
5. The module `PrintCorrectSize` appears in the Project Explorer. If a module
   with that name already exists, remove the old one first, then import again.
6. Choose **File > Save** (or Ctrl+S) to save the `.mvba` project.
7. Close the VBA editor.

If MicroStation asks whether to save the project when closing the editor,
choose **Yes**.

## 5. Load the project

If the project is not loaded:

1. Open **Utilities > Macro > Project Manager**.
2. Select `PrintCorrectSize` in the list.
3. Click **Load Project**.

The macros from the project are now available in **Utilities > Macros**.

## 6. Run the macro

You can run the macro in two ways.

### From the key-in window

Open **Utilities > Key-in** and enter:

```text
vba run [PrintCorrectSize]PrintCorrectSize.Preview
```

This only lists the models and detected page sizes. When the list looks
correct, run:

```text
vba run [PrintCorrectSize]PrintCorrectSize.Main
```

### From the Macros dialog

1. Open **Utilities > Macros**.
2. Select `PrintCorrectSize.Preview` or `PrintCorrectSize.Main`.
3. Click **Play**.

## 7. First test

Use a copy of a real DGN, not the original:

1. Open the copy.
2. Run `Preview`.
3. Check the model list and the page sizes.
4. Run `Main`.
5. Open the generated PDFs from the `PDF` sub-folder and check:
   - the page size,
   - that nothing is clipped,
   - that the drawing fills the page as expected.

The log file `PrintCorrectSize.log` is written next to the PDFs.

## Updating the macro

When a new version of `PrintCorrectSize.bas` is published:

1. Open the VBA editor.
2. Remove the old `PrintCorrectSize` module, or import the new file and let
   MicroStation replace it.
3. Import the new `.bas`.
4. Save the `.mvba` project.

The repository keeps only the current version.

## Uninstalling

1. Open the Project Manager.
2. Unload the `PrintCorrectSize` project.
3. Delete `PrintCorrectSize.mvba` from the macro folder.

The macro does not write anything outside the DGN folder and its `PDF`
sub-folder, apart from the optional log file.
