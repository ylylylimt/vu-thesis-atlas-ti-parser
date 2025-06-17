# Atlas-Ti parser 

## Before parsing:

Make sure each code value is annotated by 'Tn' (n = code number) at the end. 
Here is an example of valid project:
<img width="484" alt="Screenshot 2025-06-16 at 22 00 36" src="https://github.com/user-attachments/assets/8b71687d-ee88-41e9-bf79-d01e0a58f9c7" />

## Parsing:

- Export atlas.ti to xml 
- Run `pip install -r requirements.txt` in the terminal (ideally create a python virtual environment)
- Update `name_of_paper_xml` variable to parse the correct file
- Run main.py and open output.xlsx for results
