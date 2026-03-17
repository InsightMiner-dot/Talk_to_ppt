import os
import streamlit as st
from langchain.agents import create_agent
from langchain_openai import AzureChatOpenAI
from langchain_community.document_loaders import UnstructuredPowerPointLoader
from langchain_core.tools import tool
from langchain_core.messages import HumanMessage, AIMessage
from dotenv import load_dotenv

# --- LOAD ENV ---
load_dotenv(override=True)

# --- CONFIG FROM .env ---
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")
OPENAI_API_VERSION = os.getenv("OPENAI_API_VERSION")

os.environ["AZURE_OPENAI_API_KEY"] = AZURE_OPENAI_API_KEY
os.environ["AZURE_OPENAI_ENDPOINT"] = AZURE_OPENAI_ENDPOINT
os.environ["OPENAI_API_VERSION"] = OPENAI_API_VERSION

# --- STREAMLIT UI ---
st.set_page_config(page_title="PPT Analyst", layout="wide")
st.title("📊 PPT Analyst (Fixed Version)")

# --- TOOL ---
@tool
def analyze_slide_elements(slide_number: int = None):
    """
    Extract content from PPT. If slide_number is None, extract full PPT.
    """
    if "current_pptx_bytes" not in st.session_state:
        return "❌ No PowerPoint uploaded."

    temp_path = "active_analysis.pptx"

    # Save file safely
    with open(temp_path, "wb") as f:
        f.write(st.session_state.current_pptx_bytes)

    loader = UnstructuredPowerPointLoader(
        temp_path,
        mode="elements",
        strategy="hi_res",
        infer_table_structure=True
    )

    docs = loader.load()

    # Filter by slide if provided
    if slide_number:
        docs = [d for d in docs if d.metadata.get("page_number") == slide_number]

    if not docs:
        return f"⚠️ No content found."

    output = []
    for el in docs:
        slide_no = el.metadata.get("page_number", "?")
        category = el.metadata.get("category", "Text")
        content = el.metadata.get("text_as_html", el.page_content)

        output.append(f"[Slide {slide_no}][{category}]: {content}")

    # limit output to avoid token overflow
    return "\n".join(output[:200])


# --- LLM ---
llm = AzureChatOpenAI(
    azure_deployment=AZURE_OPENAI_DEPLOYMENT,
    temperature=0
)

# --- AGENT ---
agent = create_agent(
    model=llm,
    tools=[analyze_slide_elements],
    system_prompt=(
        "You are a PPT Analyst.\n"
        "ALWAYS use the tool 'analyze_slide_elements' before answering.\n"
        "If user does not mention slide number, analyze full PPT.\n"
        "Do NOT answer without using the tool."
    )
)

# --- SESSION STATE ---
if "messages" not in st.session_state:
    st.session_state.messages = []

# --- DISPLAY CHAT HISTORY ---
for msg in st.session_state.messages:
    role = "user" if isinstance(msg, HumanMessage) else "assistant"
    with st.chat_message(role):
        st.markdown(msg.content)

# --- CHAT INPUT ---
prompt = st.chat_input(
    "Upload PPTX or ask something...",
    accept_file=True,
    file_type=["pptx"]
)

if prompt:

    # --- HANDLE FILE UPLOAD ---
    if prompt.files:
        uploaded_file = prompt.files[0]

        file_bytes = uploaded_file.read()

        st.session_state.current_pptx_bytes = file_bytes
        st.session_state.current_pptx_name = uploaded_file.name

        msg = f"✅ Loaded file: **{uploaded_file.name}**"
        st.session_state.messages.append(AIMessage(content=msg))

        with st.chat_message("assistant"):
            st.markdown(msg)

    # --- HANDLE TEXT ---
    if prompt.text:
        user_text = prompt.text

        st.session_state.messages.append(HumanMessage(content=user_text))

        with st.chat_message("user"):
            st.markdown(user_text)

        with st.chat_message("assistant"):
            if "current_pptx_bytes" not in st.session_state:
                st.error("❌ Please upload a PPTX first.")
            else:
                with st.spinner("Analyzing PPT..."):

                    try:
                        result = agent.invoke({
                            "input": user_text
                        })

                        answer = result["output"]

                    except Exception as e:
                        answer = f"⚠️ Error: {str(e)}"

                    st.markdown(answer)
                    st.session_state.messages.append(AIMessage(content=answer))
