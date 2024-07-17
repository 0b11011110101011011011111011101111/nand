# NAND

Make your own CPU from stratch, in Excel!

This was a game I made back in 2019 to get familiar with `VBA` and in general putting together a larger program.  
It's essentially a visual remake of <a href="https://store.steampowered.com/app/576030/MHRD/"><b>MHRD</b></a>.</li>

### Installation & Startup

This file has the `.xlsm` extension, indicating that it utilizes `VBA` code (referred to as `macros` by Excel). These extensions will often trigger antivirus warnings when trying to download. I promise there's nothing suspicious about the contents of this file.

To get started, press the `NEW GAME` button from the title sheet, which will open the `Controls` panel. This panel has tabs showing us the available `LEVELS` we can play, and also which `CIRCUITS` are available to solve puzzles with. To start, we have unlocked no circuitboards. Select the first level `WIRE` to load the first level. It is recommended to always read the `LEVEL INFO` when loading a new level.

If the controls are still very unintuitive, then let's walk through some basics. Select a square of cells and press one of the 6 buttons in the `Wire Tools` section to place down some wire. Select an area and press `ERASE` to clear everything in those cells. Try to stay within the marked bounds, but if you mess things up feel free to press `RESET LEVEL`. When you unlock some additional circuits to place, first select the cell where you want to place the chip then press the corresponding button. Toggle between `Center Placement` and `Corner Placement` to determine if the selected chip is placed with the selected cell in the corner or in the center. When removing chips, please be sure to remove the entire chip and leave no "fragments" in any cells before pressing `GO!`.

### Features

- Build circuitboards to solve puzzles, using unlocked chips to recursively build more complex boards.
- Once a level is solved, the associated chip is unlocked to use in other levels.
- Non-linear level structure with password saving capability.
- Button interface via UserForm to make things easier.

### Areas of Improvement

- The circuitboard exists only on the sheet itself, without storing any other "virtual" state in memory. This requires constant interaction with the sheet object, which is not the fastest way to do things.
- Due to how the circuit flow is simulated, excessively winding wires may cause unexpected behavior.
