
import os
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from langchain_community.document_loaders import PyPDFLoader
from langchain_community.vectorstores import FAISS
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter


# Load API key from .env beside this script
load_dotenv(Path(__file__).resolve().parent / ".env")

st.set_page_config(page_title="Multi-PDF RAG App", page_icon="📚")
st.title("Multi-PDF RAG App")
st.write("Upload multiple PDFs and ask questions across them.")

if not os.getenv("OPENAI_API_KEY"):
    st.error("OPENAI_API_KEY is missing. Add it to your .env file.")
    st.stop()

# Keep the vector database during Streamlit reruns
if "store" not in st.session_state:
    st.session_state.store = None

if "indexed_files" not in st.session_state:
    st.session_state.indexed_files = []

uploaded_files = st.file_uploader(
    "Upload PDF documents",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("Build / Rebuild Index"):
    if not uploaded_files:
        st.warning("Please upload at least one PDF.")
    else:
        documents = []

        try:
            with st.spinner("Reading PDF files..."):
                for uploaded_file in uploaded_files:
                    # PyPDFLoader requires a filesystem path
                    with tempfile.NamedTemporaryFile(
                        suffix=".pdf", delete=False
                    ) as temp_file:
                        temp_file.write(uploaded_file.getvalue())
                        temp_path = temp_file.name

                    try:
                        pages = PyPDFLoader(temp_path).load()

                        for page in pages:
                            # Preserve the original uploaded filename
                            page.metadata["source"] = uploaded_file.name

                            # Convert zero-based page index to page number
                            page.metadata["page"] = (
                                int(page.metadata.get("page", 0)) + 1
                            )

                        documents.extend(pages)
                    finally:
                        Path(temp_path).unlink(missing_ok=True)

            splitter = RecursiveCharacterTextSplitter(
                chunk_size=1000,
                chunk_overlap=150,
            )
            chunks = splitter.split_documents(documents)

            if not chunks:
                st.error("No readable text was found in these PDFs.")
                st.stop()

            with st.spinner("Creating embeddings and FAISS index..."):
                embeddings = OpenAIEmbeddings(
                    model="text-embedding-3-small"
                )

                # All uploaded PDFs share one vector database
                st.session_state.store = FAISS.from_documents(
                    chunks, embeddings
                )

            st.session_state.indexed_files = [
                f.name for f in uploaded_files
            ]
            st.session_state.chunk_count = len(chunks)

            st.success(
                f"Indexed {len(chunks)} chunks from "
                f"{len(uploaded_files)} PDF file(s)."
            )

        except Exception as error:
            st.error(f"Indexing failed: {error}")

if st.session_state.store is not None:
    st.subheader("Indexed documents")
    for filename in st.session_state.indexed_files:
        st.write(f"📄 {filename}")

    st.caption(
        f"Total indexed chunks: "
        f"{st.session_state.get('chunk_count', 0)}"
    )

    st.divider()
    st.subheader("Ask a question")

    question = st.text_input(
        "Your question",
        placeholder="Ask something about your PDFs..."
    )

    if st.button("Get Answer"):
        if not question.strip():
            st.warning("Enter a question first.")
        else:
            try:
                with st.spinner("Searching documents..."):
                    retriever = st.session_state.store.as_retriever(
                        search_kwargs={"k": 5}
                    )
                    results = retriever.invoke(question)

                    if not results:
                        st.warning("No relevant information was found.")
                        st.stop()

                    context_parts = []
                    for doc in results:
                        source = doc.metadata.get("source", "Unknown")
                        page = doc.metadata.get("page", "Unknown")
                        context_parts.append(
                            f"Source: {source}, Page: {page}\n"
                            f"{doc.page_content}"
                        )

                    context = "\n\n---\n\n".join(context_parts)

                    llm = ChatOpenAI(
                        model="gpt-4o-mini",
                        temperature=0,
                    )

                    response = llm.invoke(
                        f"""Answer the question using only the
provided document excerpts. If the answer is not present,
say you could not find enough information. If sources
disagree, explain the disagreement.

Question:
{question}

Document excerpts:
{context}
"""
                    )

                st.subheader("Answer")
                st.write(response.content)

                st.subheader("Sources")
                seen = set()

                for doc in results:
                    source = doc.metadata.get("source", "Unknown")
                    page = doc.metadata.get("page", "Unknown")
                    key = (source, page)

                    if key not in seen:
                        st.caption(f"📄 {source} — Page {page}")
                        seen.add(key)

            except Exception as error:
                st.error(f"Question answering failed: {error}")
