# Excel: human label to unique code
The idea of this project is to be able to convert an excel sheet, drafted by a human being in the form of `label -> value`, to a unique coding in the form `code -> value`.

Thus the labels, which describe a certain field, must be translated into a unique code, while managing the possibility that the labels contained in the excel document will differ in syntax from the known and expected value.

To do this, a language model is used to translate the text into embeddings, and be able to perform the comparison by cosine distance and identify the best semantic proximity to which to match the unique code.

In the first instance, the reference dataset (a two-column csv file in the form `label` | `unique code`) is processed to generate the embeddings with respect to the actual, expected label and the corresponding unique code.
Then, processing one cell at a time of the excel sheet, the closest semantic similarity is sought to translate each label written in the excel into its unique code.

# dataset.csv example

| human_label |  unique_code |
| ------------- | ------------- |
| Free text for label one   | unique_code_1 |
| Free text for label two   | unique_code_2 |
| ...   | ... |

## Installation

1. **Clone the repository**:

    ```bash
    git clone https://github.com/SigfridoCorradi/excel_label_2_code
    cd excel_label_2_code
    ```

2. **Create a virtual environment** (optional but **strongly** recommended):

    ```bash
    python -m venv venv
    source venv/bin/activate
    ```

3. **Install dependencies**:

    ```bash
    pip3 install -r requirements.txt
    ```
