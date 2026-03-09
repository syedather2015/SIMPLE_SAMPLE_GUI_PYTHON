# Simple Systematic Sampler (Python GUI)

This repository contains a lightweight Python tool that performs **systematic sampling** on large Excel/CSV datasets.  
It replicates the behavior of your original Excel VBA macro, but with improved performance, automation, and ease of use.

The script opens a simple GUI where you:
1. Select the input file (Excel/CSV)
2. Optionally enter a sheet name (Excel only)
3. Enter the RPG column name (kept for compatibility; not used in sampling logic)
4. Save the sampled output file

The tool ensures:
- Ideal target sample size: **730 rows**
- Hard maximum cap: **800 rows**
- Samples are selected using **systematic sampling (every k-th row)**

Save this script as: `simple_sampler_gui.py`
