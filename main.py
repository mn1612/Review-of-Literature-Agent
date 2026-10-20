import streamlit as st
from lit_review import LitReview
import pyalex
import json
import pandas as pd

if __name__ == "__main__":
    if "api_key" not in st.session_state:
        st.session_state['api_key'] = ""
    if "email" not in st.session_state:
        st.session_state['email'] = ""
    if "topics" not in st.session_state:
        st.session_state['topics'] = dict()
    if "citations" not in st.session_state:
        st.session_state["citations"] = dict()
    if "compiled_citations" not in st.session_state:
        st.session_state["compiled_citations"] = dict()
    if "lit_reviews" not in st.session_state:
        st.session_state["lit_reviews"] = dict()

    with st.sidebar:
        st.subheader("API Key")
        api_key = st.text_input("Enter your OpenRouter API Key Here", type = "password")
        api_button = st.button("Submit", key = "input_api_key")

        if api_button:
            st.session_state.api_key = api_key
            st.write("API Key Added")

    with st.sidebar:
        st.subheader("Email")
        email = st.text_input("Enter your Email Address Here")
        email_button = st.button("Submit", key = "input_email")

        if email_button:
            st.session_state.email = email
            pyalex.config.email = email
            st.write("Email Added")

    add_topics, generate, export = st.tabs(['Add Topics/PDF Upload', 'Generate Review of Literature', 'Export Papers and Review of Literature'])
    
    with add_topics:
        st.title("Add or generate topics for literature review")
        with st.form(key = "add_topic", clear_on_submit = True):
            col_topic, col_num = st.columns([5, 1], vertical_alignment="bottom")
            with col_topic:
                new_topic = st.text_input("Enter a topic:", placeholder = "Type here")
            with col_num:
                new_num_papers = st.number_input("No. of Papers", min_value = 3, max_value = 30, value = 10)
            submit_topic = st.form_submit_button(label = "Add to List")

        if submit_topic and new_topic.strip():
            if new_topic in st.session_state.topics:
                st.error("Topic already added")
            else:
                st.session_state.topics[new_topic] = new_num_papers

        with st.form(key = "pdf_upload_form", clear_on_submit = True):
            uploaded_file = st.file_uploader("OR Choose a PDF file to automatically generate topics", type = ['pdf'])
            uploaded_file_submitted = st.form_submit_button("Process File")

            if uploaded_file_submitted and uploaded_file is not None:
                if st.session_state.api_key == "":
                    st.error("Please first add an API key")
                elif st.session_state.email == "":
                    st.error("Please first add an email")
                else:
                    if "review" not in st.session_state:
                        st.session_state["review"] = LitReview(
                            api_key = st.session_state.api_key,
                            email = st.session_state.email
                            )
                    uploaded_file_topics = st.session_state.review.get_topics(uploaded_file)
                    
                    for new_topic in json.loads(uploaded_file_topics)["topics"]:
                        st.session_state.topics[new_topic] = 10


        st.title("Current Topics")

        # 2. Loop through the items and display an interactive number field for each
        for topic in st.session_state.topics:
            
            # We create side-by-side layout for clean visuals
            col_topics_added, col_nums_added = st.columns([3, 1])
            
            with col_topics_added:
                st.markdown(f"**{topic}**")
                
            with col_nums_added:
                # Use a unique key tied to the item name so Streamlit tracks it automatically
                st.session_state.topics[topic] = st.number_input(
                    label=f"Quantity for {topic}", # Label hidden visually but required for accessibility
                    min_value=3,
                    max_value=30,
                    value= st.session_state.topics[topic],
                    key=f"input_{topic}",
                    label_visibility="collapsed" 
                )

    with generate:
        st.title("Choose topics to generate a literature review")
        selected_items = st.multiselect(
            label = "Choose topics",
            options = st.session_state.topics.keys(),
            default = st.session_state.topics.keys()
            )

        compile_literature = st.button("Generate literature reviews")
        if compile_literature:
            if st.session_state["api_key"] == "":
                st.error("Please add an API Key")
            elif st.session_state["email"] == "":
                st.error("Please add your email address")
    
            elif "review" not in st.session_state:
                st.session_state["review"] = LitReview(
                    api_key = st.session_state["api_key"],
                    email = st.session_state["email"]
                    )
            else:     
                topic_header = st.empty()
                for topic in selected_items:
                    if topic not in st.session_state.citations:
                        topic_header.subheader(f"Compiling literature review for {topic}")
                        st.session_state.citations[topic] = st.session_state.review.get_citations(
                            topic, st.session_state.topics[topic])
                        st.session_state.compiled_citations[topic] = st.session_state.review.compile_citations(
                            st.session_state.citations[topic])
                        st.session_state.lit_reviews[topic] = st.session_state.review.lit_review(
                            topic, st.session_state.compiled_citations[topic])
                    st.subheader(topic)
                    st.write(st.session_state.lit_reviews[topic])

                topic_header.subheader("Completed compiling literature reviews for selected topics")

                st.header("References")
                references = []
                for topic in selected_items:
                    for paper in st.session_state.compiled_citations[topic]:
                        references.append(st.session_state.compiled_citations[topic][paper]["apa"])
                        references.sort()
                for reference in references: st.write(reference)
                
            
    with export:
        st.title("Choose topics to export papers and literature reviews")
        selected_items_for_export = st.multiselect(
            label = "Choose topics to export",
            options = st.session_state.topics.keys(),
            default = st.session_state.topics.keys()
            )
        
        if selected_items_for_export:
            if st.session_state.compiled_citations and st.session_state.lit_reviews:

                export_papers = {
                    topic: st.session_state.compiled_citations[topic]
                    for topic in selected_items_for_export
                    if topic in st.session_state.compiled_citations
                }

                export_reviews = {
                    topic: st.session_state.lit_reviews[topic]
                    for topic in selected_items_for_export
                    if topic in st.session_state.lit_reviews
                }

                # Generate files using the selected topics
                excel_data = st.session_state.review.excel_export(
                    export_papers
                )

                word_data = st.session_state.review.word_export(
                    export_reviews, export_papers
                )

                st.download_button(
                    label="Download Excel workbook of compiled papers",
                    data=excel_data,
                    file_name="lit_review_papers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    label="Download Word document of literature reviews",
                    data=word_data,
                    file_name="lit_review.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            
            
        
            

        
