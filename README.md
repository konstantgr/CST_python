# CST_python VERSION 0
A script that allows you to simulate a grid of wires in CST Microwave Studio using frequency solver.

In file example_N3.txt stores data of (length, x, y) of wire. N3 means 3x3 grid of wires. 
In file output.txt stores postprocessing data of front scattering after plane wave excitation. 

To start using this script:
1. In <code>script.bas</code> set variable <code>myFile</code> as absolute path to input.txt on your PC.
2. In <code>main.ipynb</code> set variable <code>route</code> as absolute path to file with (length, x, y) data.
3. In <code>main.ipynb</code> set variable <code>output</code> as absolute path where you want to see an output.
4. In <code>main.ipynb</code> set variable <code>route_to_CST_DE</code> as absolute path where <i>CST DESIGN ENVIRONMENT.exe</i> stores.
5. To change parameters of frequency solver find <code>FDSolver</code> in <code>script.bas</code>. For example if you want to change number of samples change <code>.AddSampleInterval "", "", "4", "Automatic", "True" </code> attribute. 4 means number of automatical samples.
<img width="448" alt="image" src="https://user-images.githubusercontent.com/60608301/120325389-432e9d00-c2f0-11eb-9a60-60cdfc28a689.png">
