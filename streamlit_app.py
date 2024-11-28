import streamlit as st
import pandas as pd
from io import StringIO, BytesIO
from langdetect import detect, DetectorFactory
from googletrans import Translator
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from nltk.corpus import stopwords
import nltk
import re
from collections import Counter

# Ensure consistent language detection
DetectorFactory.seed = 0

# Initialize translator
translator = Translator()

# Download NLTK stopwords if not already downloaded
nltk.download('stopwords', quiet=True)

# Function to remove illegal characters
def remove_illegal_characters(text):
    """
    Removes illegal characters from a string that are not allowed in Excel cells.
    """
    # Remove characters with code points < 32 except for tab (\t), newline (\n), carriage return (\r)
    return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', text)

# Function to clean DataFrame
def clean_dataframe(df):
    """
    Cleans the DataFrame by removing illegal characters from all string-type columns.
    """
    string_cols = df.select_dtypes(include=['object']).columns
    for col in string_cols:
        df.loc[:, col] = df[col].apply(lambda x: remove_illegal_characters(str(x)) if pd.notnull(x) else x)
    return df

# Function to load data
@st.cache_data
def load_data(uploaded_file, pasted_data):
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            st.error("Unsupported file format!")
            return None
    elif pasted_data:
        try:
            df = pd.read_csv(StringIO(pasted_data), sep=None, engine='python')
        except Exception as e:
            st.error(f"Error parsing pasted data: {e}")
            return None
    else:
        df = None
    return df

# Function to detect language
def detect_language(text_series):
    try:
        sample_text = ' '.join(text_series.dropna().astype(str).tolist()[:100])  # Use first 100 entries for detection
        lang = detect(sample_text)
    except Exception:
        lang = 'unknown'
    return lang

# Function to translate text
def translate_text(text, src, dest, retries=3, delay=5):
    for attempt in range(retries):
        try:
            translated = translator.translate(text, src=src, dest=dest)
            return translated.text
        except Exception as e:
            if attempt < retries - 1:
                st.warning(f"Translation failed for '{text}'. Retrying in {delay} seconds...")
                time.sleep(delay)
            else:
                st.error(f"Translation error for '{text}': {e}")
                return text

# Streamlit App
def main():
    st.set_page_config(page_title="Excel Translator & Word Analyzer", layout="wide")
    st.title("📊 Excel Translator & Word Analyzer")

    # Tabs
    tabs = st.tabs(["🔄 Translate", "📝 Word Analysis"])

    with tabs[0]:
        st.header("Translation Module")

        # File uploader and text area
        uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["xlsx", "xls", "csv"])
        st.write("OR")
        pasted_data = st.text_area("Paste your CSV data here")

        df = load_data(uploaded_file, pasted_data)

        if df is not None:
            st.success("Data loaded successfully!")
            st.dataframe(df.head())

            # Select column to translate
            columns = df.columns.tolist()
            column_to_translate = st.selectbox("Select the column to translate", columns)

            if column_to_translate:
                texts = df[column_to_translate].dropna().astype(str)

                # Detect language
                lang = detect_language(texts)
                st.write(f"**Detected Language:** {lang}")

                if lang == 'unknown':
                    st.error("Could not detect language. Please ensure the text is sufficient for detection.")
                else:
                    # Select target language
                    LANGUAGES = {
                        'English': 'en',
                        'Spanish': 'es',
                        'French': 'fr',
                        'German': 'de',
                        'Chinese (Simplified)': 'zh-cn',
                        'Japanese': 'ja',
                        'Arabic': 'ar',
                        'Hindi': 'hi',
                        'Portuguese': 'pt',
                        'Russian': 'ru',
                        'Italian': 'it'  # Added Italian
                    }

                    target_lang_name = st.selectbox("Select target language", list(LANGUAGES.keys()), index=0)
                    target_lang = LANGUAGES[target_lang_name]

                    if st.button("Translate"):
                        with st.spinner("Translating..."):
                            translated_texts = texts.apply(lambda x: translate_text(x, src=lang, dest=target_lang))
                            df[f"{column_to_translate}_translated"] = translated_texts
                        st.success("Translation completed!")
                        st.dataframe(df[[column_to_translate, f"{column_to_translate}_translated"]].head())

                        # Option to download the translated data
                        to_download = df[[column_to_translate, f"{column_to_translate}_translated"]]
                        to_download = clean_dataframe(to_download)  # Clean the DataFrame
                        to_download_buffer = BytesIO()
                        to_download.to_excel(to_download_buffer, index=False, engine='openpyxl')
                        to_download_bytes = to_download_buffer.getvalue()

                        st.download_button(
                            label="Download Translated Data as Excel",
                            data=to_download_bytes,
                            file_name='translated_data.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )

    with tabs[1]:
        st.header("Word Analysis Module")

        if df is not None:
            # Select column for analysis
            analysis_column = st.selectbox("Select the column for word analysis", df.columns.tolist(), key="analysis_column")

            if analysis_column:
                texts = df[analysis_column].dropna().astype(str).tolist()
                combined_text = ' '.join(texts)

                # Detect language for stopwords
                lang = detect_language(df[analysis_column])

                # Get stopwords
                if lang in stopwords.fileids():
                    lang_stopwords = set(stopwords.words(lang))
                else:
                    lang_stopwords = set()

                remove_sw = st.checkbox("Remove Stopwords", value=True)

                if remove_sw and lang_stopwords:
                    words = combined_text.split()
                    filtered_words = [word for word in words if word.lower() not in lang_stopwords]
                    final_text = ' '.join(filtered_words)
                else:
                    final_text = combined_text

                if final_text.strip():  # Check if final_text is not empty or just whitespace
                    # Word Cloud
                    st.subheader("Word Cloud")
                    wordcloud = WordCloud(width=800, height=400, background_color='white').generate(final_text)
                    fig_wc, ax_wc = plt.subplots(figsize=(10, 5))
                    ax_wc.imshow(wordcloud, interpolation='bilinear')
                    ax_wc.axis('off')
                    st.pyplot(fig_wc)

                    # Word Frequency
                    st.subheader("Word Frequency")
                    word_counts = Counter(final_text.split())
                    most_common = word_counts.most_common(20)
                    freq_df = pd.DataFrame(most_common, columns=['Word', 'Frequency'])
                    st.dataframe(freq_df)

                    # Bar Chart for Word Frequency
                    st.subheader("Word Frequency Chart")
                    fig_freq, ax_freq = plt.subplots(figsize=(10, 5))
                    ax_freq.bar([x[0] for x in most_common], [x[1] for x in most_common], color='skyblue')
                    ax_freq.set_xlabel('Words')
                    ax_freq.set_ylabel('Frequency')
                    ax_freq.set_title('Top 20 Words')
                    plt.xticks(rotation=45)
                    st.pyplot(fig_freq)
                else:
                    st.warning("No words available for analysis. Please ensure the selected column contains valid text and that stopwords removal did not eliminate all words.")
        else:
            st.warning("Please load data in the Translate tab first.")

if __name__ == "__main__":
    main()
