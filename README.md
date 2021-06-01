# CST_python VERSION 0
A script that allows you to simulate a grid of wires in CST Microwave Studio using frequency solver.

In file example_N3.txt stores data of (length, x, y) of wire. N3 means 3x3 grid of wires. 
In file output.txt stores postprocessing data of front scattering after plane wave excitation. 

To start using this script:
1. In <code>script.bas</code> set variable <code>myFile</code> as absolute path to input.txt on your PC.
2. In <code>main.ipynb</code> set variable <code>route</code> as absolute path to file with (length, x, y) data.
3. In <code>main.ipynb</code> set variable <code>output</code> as absolute path where you want to see an output.
