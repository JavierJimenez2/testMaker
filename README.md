# TestMaker
## Overview

The TestMaker is a Python-based tool that generates a custom multiple-choice exam based on questions from an Excel file. The script selects questions randomly, presents them to the user for answering, tracks correct and incorrect responses, and generates a report with the results, including score, questions answered correctly, and incorrect answers.

# Galería de Imágenes

| Imagen 1 | Imagen 2 | Imagen 3 |
|----------|----------|----------|
| ![Imagen 1](https://github.com/JavierJimenez2/testMaker/blob/master/Testmaker.png?raw=true) | ![Imagen 2](https://github.com/JavierJimenez2/testMaker/blob/master/Testmaker2.png?raw=true) | ![Imagen 3](https://github.com/JavierJimenez2/testMaker/blob/master/Testmaker3.png?raw=true) |


## Features

- Selects questions randomly from an Excel workbook.
- Option to filter questions by specific topics and sheets.
- Supports true/false type questions with optional images.
- Keeps track of user responses, showing feedback and justifications for correct answers.
- Generates a final score and stores the results in an Excel file.
- Option to retake the exam.

## Dependencies

- `openpyxl` - For working with Excel files.
- `PIL` (Pillow) - For handling images.
- `datetime` - For handling time and dates.
- `os` - For file and system operations.

### Installation

Before running the script, ensure you have all the required dependencies installed.

You can install the required libraries via pip:

```bash
pip install openpyxl pillow
```

## How It Works

1. **Input File (Excel):**
   - The script requires an Excel file (`.xlsx`) where each sheet represents a set of questions.
   - Each question is structured in columns as follows:
     - Column 1: The question text.
     - Column 2: The correct answer (either "V" or "F").
     - Column 3: An image path (optional).
     - Column 4: The question iteration count.
     - Column 5: The year (optional).
     - Column 6: The justification for the correct answer (optional).

2. **User Interaction:**
   - The user is prompted to select a topic for the exam (e.g., `7`, `8`, `9`, or `lab`).
   - They can select specific content or ask for a random selection.
   - The script then presents the questions one by one and asks the user to input their answer ("V" for True or "F" for False).

3. **Image Handling:**
   - If a question includes an image (referenced in Column 3), it is displayed using the default image viewer on the system.
   
4. **Feedback:**
   - After each question, the user receives feedback on whether their answer is correct or not. If correct, they are shown the justification (if provided).
   
5. **Results:**
   - The final score is calculated based on correct answers, with incorrect answers reducing the score. The results are saved in an Excel file (`resultados.xlsx`) for future reference.

6. **Retake Option:**
   - At the end of the exam, the user is asked if they want to retake the exam. If yes, the exam is reset and they can start again.

## Configuration and Usage

### 1. Start the script

To begin, run the script in your terminal or command line:

```bash
python testmaker.py
```

The script will first ask for the exam topic:

- For unit 7: Enter `7`
- For unit 8: Enter `8`
- For unit 9: Enter `9`
- For a practice exam (`lab`): Enter `l`
- For a review exam (repetition of random questions from all units): Enter `r`

### 2. Choose Specific Content (Optional)

You can choose specific sheets within a workbook by entering "s" when prompted for specific content. The script will list available sheets and allow you to select from them.

### 3. Enter Number of Questions

You will then be prompted to specify how many questions you want in the exam. If left blank, it defaults to the maximum number of questions available.

### 4. Answering the Questions

For each question, you will be shown:

- The question text.
- An optional image (if available).
- You need to input your answer: "V" for True or "F" for False.

### 5. Results

After the exam, the script will calculate your score based on the number of correct and incorrect answers. The score will be displayed, and the results will be saved to `resultados.xlsx`.

### 6. Retake Option

After completing the exam, you will be asked if you want to retake it. If you want to take another exam, simply type "s" and the exam will restart.

## File Structure

```
TestMaker/
├── testmaker.py           # Main Python script
├── docs/                  # Directory containing the Excel files (e.g., u7.xlsx, u8.xlsx, etc.)
├── img/                   # Directory containing images for questions (optional)
├── resultados.xlsx        # File where the results are saved
```

## Example Excel File Format

The Excel file used for generating the exam should have the following structure:

| Question               | Answer | Image      | Iteration | Year | Justification |
|------------------------|--------|------------|-----------|------|---------------|
| What is 2+2?            | V      | img/2plus2.jpg | 0         | 2024 | 2+2 equals 4. |
| The Earth is flat.      | F      |            | 1         |      | The Earth is round. |

## Example Output

Here’s an example of what you would see during the exam:

```
Examen tipo test tema 7:
-----------------
Aciertos: 0/5  |  Fallos: 0/5
-------------------------------------------
Pregunta([Unidad 7][Año: 2024] - Hoja 1) [1 de 5]:
¿Cuál es la suma de 2+2?
V o F: V
¡Respuesta correcta!

Pregunta([Unidad 7][Año: 2024] - Hoja 1) [2 de 5]:
La Tierra es plana.
V o F: F
Respuesta incorrecta.
Respuesta correcta: F
Justificación: La Tierra es redonda.
```

## Results Saving

The results of each exam are saved to `resultados.xlsx` with the following columns:

- **Tema**: The topic of the exam (e.g., 7, 8, 9, lab).
- **Apartado**: The list of sections used for the exam.
- **Duración**: The time taken to complete the exam.
- **Fecha**: The date and time the exam was taken.
- **Preguntas**: Total number of questions.
- **Aciertos**: Number of correct answers.
- **Fallos**: Number of incorrect answers.
- **Nota**: The final score.
