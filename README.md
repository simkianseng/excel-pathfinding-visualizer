# Excel-VBA-Interactive-Pathfinding-Visualizer

Just a chemical engineering student who is intrested field of computer science. I noticed that not many pathfinding visualization programs are written in VBA so it would be fun to create one. Just a personal project of mine, code here may be unoptimized, hope you enjoy playing around with the program.

## Application Features
1. Interactive user interface
2. Unlimited map area (Excel maximum is 1,048,576 rows by 16,384 columns)
3. Wide range of maze editing tools:
    1. **Eraser** - Allows user to remove selected obstacles/ start point/ end point placed down on the cells.
    2. **Place obstacle** - Places an obstacle on selected cells that obstructs the path (cannot be explored).
    3. **Place start point** - Places a start point on a selected cell.
    4. **Place end point** - Places an end point on a selected cell.
4. Different search algorithms:
    1. **Breadth First Search** - An alogrithm that finds the shortest path from the start to end.
    2. **Depth First Search** - An algorithm that finds a path from the start to end point.
    3. **Greedy Best First Search** - A fast algorithm that finds a path from the start to end point.
    4. **A Star Search** - A fast algorithm that finds the shortest path from the start to end point.
5. Random maze generator
6. Animation settings
    1. **Show explored cells** - A setting that can be toggled on or off depending if the user wants to see the cells explored by the algorithms.
    2. **Checkbox to show actual path** - A setting that can be toggled on or off depending if the user wants to see the actual path found by the algorithm.
    3. **Show explored cells delay time** - A setting that user can change the value of, affecting the animation speed of the explored cells appearing.
    4. **Show actual path delay time** - A setting that user can change the value of, affecting the animation speed of the actual path appearing.


## Credits
1. Breath First Search, Depth First Search and A* Search algorithms used are in Python code translated to VBA and modified from: https://www.linkedin.com/learning/python-data-structures-and-algorithms/python-data-structures-and-algorithms-in-action?u=76881922
2. Code for random maze generator is taken and modified from: https://github.com/stelios7/ExcelPathfinding
3. Project is inspired from: https://www.youtube.com/watch?v=1umm4PvD8n0
