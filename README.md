# minesweeper-with-hancell
Anything is possible at the Ministry of National Defense...
- But due to restrictions of Hancell, disabling the right-click menu via `Cancel = True` is not possible...
  - so the plan is to provide open mode/flag mode seperately, just as in most mobile minesweeper apps.

## Current Implementation Status
### 2026-06-03: First Commit
- Start(not play) new game
- Internal initialization
- Random mine placement
- Calculation of surrounding mine counts for each cell
- Display of all cells in the closed state
- `RenderBoard` function to reveal the full board for debugging
### 2026-06-03: Second Commit
- Link cell click events
- Add open mode, flag mode
- Open or flag cell when clicked
- Trigger game over when a mine is opened

 **(Barely) Playable from here!**


## Current Module Structure
- `modConfig`: Configuration values (e.g. board size and mine count)
- `modState`: Game state array management
- `modGame`: Overall game flow, including starting a new game
- `modBoard`: Mine placement and surrounding mine count calculation
- `modRender`: Cell rendering

## Remaining Tasks
- Add cell reveal logic for "openings"
- Add win condition checking
- Complete game status logic(ready, ongoing, win, game over)
- First click should not be a mine
- Add timer
- Edit gameboard and buttons layout & design
  - Place images for mines and flags instead of "M" and "F"
  - Color numbers and use 8-bit font if possible
- Add difficulty choices(beginner, intermediate, expert) and custom board sizes
- Change smiley face expressions according to game status
- Add No Guess Mode