import openai
from docx import Document
import os
from config import OPENAI_API_KEY

openai_client = openai.OpenAI(api_key=OPENAI_API_KEY)


def analyze_document():
    file_path = "raw_data_thomas_consulting.docx"
    company_name = "Thomas Consulting Group"

    try:
        document = Document(file_path)
        full_text = "\n".join([paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()])
        if not full_text:
            return "Error: The document is empty or could not be read."

    except Exception as e:
        return f"Error reading the document: {e}"

    prompt = f"You are an expert financial analyst tasked with evaluating potential investors for {company_name}. " \
             f"Below is a document listing companies that may invest in {company_name}, along with relevant details " \
             f"about their investment interests and preferences. " \
             f"Analyze the document carefully and generate a ranked list of all these companies, from most likely to " \
             f"least likely to invest in {company_name}. " \
             f"Do not base the ranking on the order in which the companies are listed. Instead, consider factors such " \
             f"as their stated interest in consulting firms as well as strategic alignment with the company, " \
             f"and any other relevant details provided. " \
             f"If critical information is missing or ambiguous, make reasonable assumptions and explain your " \
             f"reasoning briefly in the response. " \
             f"Provide a clear, ordered list with a short justification for each company’s position as well as a " \
             f"description of the company and keep them concise. \n\nDocument " \
             f"content:\n{full_text}"
    try:
        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=1500,
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error analyzing filings with OpenAI: {e}"


def main():
    result = analyze_document()
    print(result)


if __name__ == "__main__":
    main()
