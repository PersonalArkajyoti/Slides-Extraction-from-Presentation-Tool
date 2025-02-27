# Deckorator
  Deckorator is a proof-of-concept (PoC) tool that allows users to upload and process presentation files. It utilizes keyword-based search functionality to identify and extract relevant slides from the uploaded decks.
  ## Features

    1. Upload any presentation file and search for content using embedded keywords.
    2. Keyword-based filtering: The tool identifies slides based on keywords that are embedded in white text within the deck.
    3. Selection and extraction: Users can select relevant slides and download them as a new presentation.
    4. Automated processing: The system processes slides efficiently to deliver accurate search results.
    5. Watermark removal: Ensures clean and professional output by eliminating unwanted watermarks.

  ## How It Works

    - Upload a presentation file.
    - Use the search bar to find relevant slides based on pre-embedded keywords.
    - Select slides from the matching results.
    - Download the selected slides as a new presentation file.

  ## Installation & Usage

  ### Clone this repository:

    git clone https://github.com/your-username/deckorator-v2.1.git
    cd deckorator-v2.1

  ### Install dependencies:

    pip install -r requirements.txt

  ###  Run the application:

    streamlit run app.py
  Upload a presentation file and search for relevant slides using keywords.

  ## Requirements

    - Python 3.8+
    - Streamlit
    - Aspose.Slides
    - PIL (Pillow)
  Other dependencies listed in requirements.txt

  ## Notes

    - Ensure that keywords are embedded in white text inside the presentation files.
    - This project is a proof-of-concept and may require further refinement for production use.

  ## License

    This project is licensed under the MIT License.

  ## Contributing
    Contributions are welcome! Feel free to fork this repository and submit a pull request with your improvements.

  ## Contact

  For any queries or suggestions, feel free to open an issue or reach out at arkajyotichakraborty99@gmail.com
