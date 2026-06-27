# minesweeper-with-hancell

<img width="853" height="535" alt="image" src="https://github.com/user-attachments/assets/73085230-4933-4162-a68c-9011fb60820d" />

Anything is possible at the Ministry of National Defense...
- But due to restrictions of Hancell, disabling the right-click menu via `Cancel = True` is not possible...
  - so the plan is to provide open mode/flag mode seperately, just as in most mobile minesweeper apps.
 
**The .xlsm file only opens on Hancell, and does not work on Excel.**

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

### 2026-06-05
- Add win condition checking

**Truly playable from here!** (but still extremely unfriendly)
- Disable range selection to prevent confusion

### 2026-06-06
- Add cell reveal logic for "openings"
- Add colors to mine count numbers(1: blue, 2: green, ...)
- Show unicode flag and mine instead of "F" and "M"
- Complete game status logic(ready, ongoing, win, game over)
- First click is not a mine
  - Game starts and places mines after player's first click
- Add (total mine count - flag count)
- Show all mines at the end of a game(flags for a win, mines for a loss)
  - Triggered mine is painted red, wrong flags are painted pink
- Add timer
  - Down to seconds while playing, down to thousandths after the game

### 2026-06-07
- Restyle game start button(smiley face)
- Change smiley face expressions according to game status
- Round the elapsed time for better experience
- Add chording logic
  - Implement game over by chording with wrong flag
- Fix issue: flagged cells not (internally) unflagging when opened by "opening" expansion

**Fully functioning game from here**

### 2026-06-08
- Resize the board to Expert(30x16/99)
- Optimize rendering for larger boards
  - Reduce unneccesary access, redundant rendering to various things

### 2026-06-13
- Place minesweeper-style images for mines, flags, numbers

### 2026-06-14
- Place digital number images for timer, mine count

### 2026-06-15
- Place number background image

## Current Module Structure
- `modConfig`: Configuration values (e.g. board size and mine count)
- `modState`: Game state array management
- `modGame`: Overall game flow, including opening/flagging cells
- `modBoard`: Mine placement and surrounding mine count calculation
- `modRender`: Cell rendering
- `modTimer` : Timer control
- `modUtils` : Utility functions/subs, Test macros

## Remaining Tasks
### Soon
- Fix: Wrong flag being removed when cell opened via openings by chording
- Fix: Disable shape selection with right click
- Redesign mode buttons
- Add "pressed" image for better experience
- Add difficulty choices(beginner, intermediate, expert) and custom boards
- Show detailed game info after the game
  - exact time, (statistics)
### Later, if not never (in likely-to-be-done order)
- Add Ranking
- Add statistics(3BV, Efficiency, etc.)
- Add Hint
- Add No Guess Mode
