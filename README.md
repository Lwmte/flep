FLEP (FLExible Patcher) aka TREP2
---------------------------------

### Table of contents ###

1.  What is this?
2.  Features
3.  What's new?


1. What is this?
----------------

FLEP is a multi-purpose binary patcher, which can be used to create, manage and apply different
binary patches to different files. Primarily, it was developed for Tomb Raider community to operate
with so-called *TRNG* binary file (tomb4.exe, which is used to play custom Tomb Raider levels built
with patched level editor).

FLEP's source code is primarily based on old TRLE patcher utility called *TREP*, which was done in
VB6. Generally, code was refactored, cleaned up, some modern approaches and UI tricks were used to
make it look more modern and less lame.


2. Features
-----------

* 1) Almost completely rewritten and refactored source code. As result, application works much faster.
* 2) New interpreter of static patches. It now supports not only raw hex strings, but also some different commands which alter patching behaviour - SetFileLength and Fill.
* 3) Allows to simultaneously patch numerous binaries.
* 4) Supports additional parameter types - strings, float, 8/16 bits and A,B,C byte sequences (useful for RGB values).
* 5) Supports unlimited amount of patches, also unlimited amount of offsets and parameters for each.
* 6) Each parameter can now have several offsets, this way you can simultaneously change similar values in binary with just only one parameter slot.
* 7) Each parameter now has its own conditional behaviour switch.


3. What's new?
--------------

* 1.1.41  - RGB data type modified to support scattered offsets for each colour byte (divide them with "|" symbol).

* 1.1.38  - Minor code refactoring, fixed wrong patch set / preset handling.

* 1.1.31  - Quick overflow fix for last bit of bits(16) datatype.

* 1.1.30  - Hex types replaced with new bit types, modified parameter view a bit to support multi-column mode, added copy function, minor interface fixes and speed-ups.

* 1.0.86  - Fixed param list killing with escape button, re-done patching method with kernel32 routines (faster), fixed various bugs with moving patch up/down, added minimize button, added color picker for RGB param.

* 1.00b   - Beta version. Fixed various UI bugs, changed preset format a bit.

* 1.00a   - Initial alpha version.
