# minesweeper-with-hancell
Anything is possible at the Ministry of National Defense...
- But due to restrictions of Hancell, disabling the right-click menu via `Cancel = True` is not possible...
  - so the plan is to provide open mode/flag mode seperately, just as in most mobile minesweeper apps.

## Current Implementation Status
- Start(not play) new game
- Internal initialization
- Random mine placement
- Calculation of surrounding mine counts for each cell
- Display of all cells in the closed state
- `RenderBoard` function to reveal the full board for debugging

## Current Module Structure
- `modConfig`: Configuration values (e.g. board size and mine count)
- `modState`: Game state array management
- `modGame`: Overall game flow, including starting a new game
- `modBoard`: Mine placement and surrounding mine count calculation
- `modRender`: Cell rendering

## Remaining Tasks
- Link cell click events
- Add open mode, flag mode
- Open or flag cell when clicked
- Trigger game over when a mine is opened
- Add cell reveal logic for "openings"
- Add win condition checking
