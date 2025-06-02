import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from urllib.parse import urlparse
from together import Together
from pygooglenews import GoogleNews
import io
import re  # For regular expressions to extract AI reason
import xlsxwriter  # Required for writing multiple sheets to Excel

# --- Configuration ---
# Replace with your actual API key for Together AI
together_client = Together(
    api_key="a510758b9bff7bf393548b99848a45972486dd1d699eb86a5e7735d2339c1d8c"
)

# --- Prompts ---
PROMPT_INDIVIDUAL_ANALYSIS = """Carefully analyze the following news article text for information directly indicating potential reasons for client churn specifically for an **employee benefits company in India**. Focus only on details that would impact an employee benefits provider or suggest a company might reduce or discontinue its employee benefits programs.

**Text:**
{provided_text}

Based on your analysis and using the provided categories below, determine the churn risk and the specific reason(s).

1.  **Risk Level (First Line):** State the risk level as one of the following:
    * "High Risk"
    * "Medium Risk"
    * "Low Risk"
    * "No Churn Risk Indicated" (If no relevant information is found regarding churn for an employee benefits company)

2.  **Reason(s) for Risk (Second Line):** If a risk is indicated, explain the major reason(s) concisely, referencing the relevant category (e.g., "Reason: [Category Name] - Brief explanation."). If there are multiple relevant reasons, list them clearly.

3.  **2-Line Summary of Analysis (Third and Fourth Lines):** Provide a brief, overall summary of the article's relevance to churn for an employee benefits company, condensing the key findings into exactly two lines. If no churn risk is indicated, summarize why the article is not relevant.

**Categories for Reasons:**
I. Corporate Restructuring (Mergers, Acquisitions, Joint Ventures, IPO, Entity Realignment, Rebranding, Consolidation, Subsidiary changes)
II. Business Discontinuity (Closures, Market Exits, Bankruptcy, Operational Suspensions, Business Model Pivots)
III. Strategic Policy Changes (Benefits Strategy Transformation, Leadership Changes impacting strategy, Cost Optimization related to benefits, Changes in top leadership impacting benefits)
IV. Financial Constraints (Cash Flow Issues, Cost-Cutting impacting benefits, Budget Reallocation away from benefits, Severe financial loss)
V. Employment Structure Changes (Workforce Reorganization, Shifts to contractual work, Remote work transitions impacting benefits, Layoffs, Furloughs, Downsizing)
VI. Regulatory & Compliance Factors (India Specific: Changes in tax policy, GST, labor codes, social security impacting benefits compliance or costs)
VII. Competitive Market Dynamics (Client switched vendor, New platform adoption by client, Competitor activity in benefits space, Pricing pressures on benefits, Market share shifts impacting client's ability to offer benefits, Disruption in client's industry affecting benefits, Client's value proposition change impacting benefits)
VIII. Technological Transitions (Digital transformation affecting benefits administration, HRMS integration impacting benefits systems, API changes relevant to benefits platforms, Analytics adoption impacting benefits, Mobile app for benefits, Platform upgrade for benefits management)
IX. Service Delivery Issues (Onboarding delay with benefits provider, Tech issues with benefits platform, Merchant issue impacting benefits, Support problem with benefits services, Delivery delay of benefits, Reimbursement issue with benefits claims)
X. Employee Engagement (Low adoption of benefits programs, Poor user experience with benefits platform, Negative employee feedback on benefits, Generation gap affecting benefits appeal, Hybrid work models impacting benefits usage, Usage drop in benefits offerings)

**Example Output Format (for High/Medium/Low Risk):**
High Risk
Reason: Business Discontinuity - Company announced complete shutdown impacting all operations including benefits.
Summary: The company is facing imminent closure, directly impacting its ability to retain any employee benefits plans. This represents a critical churn event for any associated benefits provider.

**Example Output Format (for No Risk):**
No Churn Risk Indicated
Summary: The article discusses general market trends not specific to the company's operational or financial health. It provides no indication of changes relevant to employee benefits or potential churn.
"""

PROMPT_COMBINED_ANALYSIS = """Given the individual analyses of news articles related to a company and potential client churn, provide an overall summary (at most 4 lines).

**Individual Article Analyses:**
{individual_analyses_summary}

Prioritize information from articles that indicate a "High Risk", "Medium Risk", or "Low Risk" of churn. Only if ALL individual analyses clearly indicate "No Churn Risk Indicated" should your overall conclusion be "Overall No Churn Risk Indicated".

In the first line, state the overall risk level for churn for the company (e.g., "Overall High Risk," "Overall Medium Risk," "Overall Low Risk," "Overall No Churn Risk Indicated").
In the second line, provide the major reason(s) for this overall risk, drawing from the categories mentioned in the individual analyses, formatted as "Reason: [Category Name] - Brief explanation.".
In the subsequent lines, concisely summarize the key findings. If no relevant information is found across all articles (i.e., all articles were "No Churn Risk Indicated"), then for the reason, state "Reason: Not Applicable - No relevant churn indicators found."

**Example Overall Summary Output (for risk):**
Overall High Risk
Reason: Business Discontinuity - Company announced complete shutdown.
The company is facing imminent closure, directly impacting its ability to retain any employee benefits plans.
This represents a critical churn event for any associated benefits provider.

**Example Overall Summary Output (for no risk):**
Overall No Churn Risk Indicated
Reason: Not Applicable - No relevant churn indicators found.
All analyzed articles did not indicate any direct churn risk for the company. The news was either irrelevant or positive.
"""

PROMPT_REASON_COMPARISON = """You are an expert in client churn analysis for employee benefits companies. Your task is to compare two churn reasons and determine if they are a "Match", "Partial Match", or "No Match".

**Reason from AI News Analysis:**
{ai_reason}

**Provided Churn Reason (from Client Data):**
{excel_reason}

Consider the meaning and intent behind both reasons.
- A "Match" means the reasons are highly similar or convey the same core cause for churn.
- A "Partial Match" means there is some overlap or a connection between the reasons, but they are not identical.
- A "No Match" means the reasons are unrelated or contradictory.

Provide exactly only one of the following as your answer: "Match", "Partial Match", or "No Match".
"""

# --- Functions ---


# Cache results for 1 hour to avoid repeated API calls
@st.cache_data(ttl=3600)
def analyze_text(company_name, provided_text, prompt_template, _together_client):
    """Analyzes the provided text for churn indicators or performs comparison using Together AI."""
    # The provided_text is already the formatted prompt content from the calling function.
    try:
        response = _together_client.chat.completions.create(
            model="meta-llama/Llama-3.3-70B-Instruct-Turbo-Free",
            messages=[{"role": "user", "content": provided_text}]
        )
        output = response.choices[0].message.content
        return output if output else f"Unexpected empty response for {company_name}"
    except Exception as e:
        st.error(f"Error querying Together AI for {company_name}: {e}")
        return "Analysis failed due to AI service error."

# Specific function for LLM reason comparison, separate from general text analysis


def get_llm_reason_comparison_result(ai_reason: str, excel_reason: str, _together_client) -> str:
    """
    Uses the LLM to compare two churn reasons and return "Match", "Partial Match", or "No Match".
    """
    comparison_prompt_content = PROMPT_REASON_COMPARISON.format(
        ai_reason=ai_reason, excel_reason=excel_reason)

    try:
        response = _together_client.chat.completions.create(
            model="meta-llama/Llama-3.3-70B-Instruct-Turbo-Free",
            messages=[{"role": "user", "content": comparison_prompt_content}]
        )
        llm_output = response.choices[0].message.content.strip()
        # Ensure the output is one of the expected categories
        if llm_output in ["Match", "Partial Match", "No Match"]:
            return llm_output
        else:
            st.warning(
                f"LLM returned unexpected comparison result: '{llm_output}'. Defaulting to 'No Match'.")
            return "No Match"
    except Exception as e:
        st.error(f"Error calling LLM for reason comparison: {e}")
        return "No Match (LLM Error)"


@st.cache_data(ttl=3600)  # Cache news fetching for 1 hour
def fetch_news(company_name, from_date, to_date, max_articles=10, queries=None, allowed_domains=None):
    """
    Fetches news articles for a given company using the pygooglenews library.
    Filters articles by allowed domains.
    """
    gn = GoogleNews(lang='en', country='IN')
    results = []
    if queries is None:
        queries = [company_name]

    try:
        # Process queries in groups of 3 to optimize API calls
        for i in range(0, len(queries), 3):
            group_queries = queries[i:i+3]
            combined_query = " OR ".join(group_queries)
            from_date_str = from_date.strftime('%Y-%m-%d')
            to_date_str = to_date.strftime('%Y-%m-%d')
            search_results = gn.search(
                combined_query, from_=from_date_str, to_=to_date_str)

            if search_results and 'entries' in search_results:
                articles_for_query = []
                if allowed_domains:
                    for article in search_results['entries']:
                        source_link = article.get('source', {}).get('href', '')
                        parsed_uri = urlparse(source_link)
                        domain = parsed_uri.netloc.replace('www.', '')
                        if any(d in domain for d in allowed_domains):
                            articles_for_query.append(article)
                    # If no articles from allowed domains were found, add the top article as a fallback
                    if not articles_for_query and search_results['entries']:
                        articles_for_query.append(search_results['entries'][0])
                else:
                    articles_for_query = search_results['entries']

                results.extend(articles_for_query[:max_articles])
            else:
                st.warning(
                    f"No results or 'entries' not found for query '{combined_query}'")
    except Exception as e:
        st.error(f"Error fetching news for {company_name}: {e}")
        return None
    # Ensure total articles returned is at most max_articles
    return results[:max_articles]


def process_article(article):
    """Extracts summary or title from a news article."""
    return article.get('summary') or article.get('title') or ""


def analyze_news(company_name, from_date, to_date, max_articles=10, queries=None, allowed_domains=None):
    """
    Fetches news articles for a company and analyzes them for churn indicators.
    """
    st.subheader(f"Analyzing News for **{company_name}**")
    all_articles = fetch_news(company_name, from_date,
                              to_date, max_articles, queries, allowed_domains)

    if not all_articles:
        return {"individual_analyses": [], "overall_summary": "Overall No Churn Risk Indicated\nReason: Not Applicable - No relevant churn indicators found.\nSummary: No news articles found for analysis."}

    individual_analyses_list = []
    combined_analysis_text_for_model = ""

    for i, article in enumerate(all_articles):
        article_text = process_article(article)
        article_url = article.get('link', 'No URL available')
        article_title = article.get('title', f"Article {i+1}")

        if article_text:
            # Format prompt for individual analysis
            individual_analysis_prompt_content = PROMPT_INDIVIDUAL_ANALYSIS.format(
                provided_text=article_text)  # Removed company_name from format as prompt doesn't use it
            analysis_result = analyze_text(
                company_name, individual_analysis_prompt_content, PROMPT_INDIVIDUAL_ANALYSIS, together_client)
            individual_analyses_list.append({
                "title": article_title,
                "url": article_url,
                "analysis": analysis_result
            })
            combined_analysis_text_for_model += f"Article {i+1} Analysis:\n{analysis_result}\n\n"
        else:
            no_text_analysis = "No Churn Risk Indicated\nReason: Not Applicable - No text in article summary/title.\nSummary: This article provided no relevant content for analysis."
            individual_analyses_list.append({
                "title": article_title,
                "url": article_url,
                "analysis": no_text_analysis
            })
            combined_analysis_text_for_model += f"Article {i+1} Analysis:\n{no_text_analysis}\n\n"

    # Determine if any articles indicated risk
    any_risk_found = any(get_risk_level_from_text(a['analysis']) not in [
                         "No Churn Risk Indicated", "Unknown Risk"] for a in individual_analyses_list)

    overall_summary_result = ""
    if individual_analyses_list:
        combined_prompt_content = PROMPT_COMBINED_ANALYSIS.format(
            individual_analyses_summary=combined_analysis_text_for_model.strip())

        # If no risk was found in any individual article, explicitly tell the LLM to conclude no risk.
        # Otherwise, let the LLM synthesize from risk-indicating articles.
        if not any_risk_found:
            explicit_no_risk_prompt = f"{combined_prompt_content}\n\nIMPORTANT: All individual analyses indicate 'No Churn Risk Indicated'. Therefore, your overall summary MUST be 'Overall No Churn Risk Indicated'."
            overall_summary_result = analyze_text(
                company_name, explicit_no_risk_prompt, PROMPT_COMBINED_ANALYSIS, together_client)
        else:
            overall_summary_result = analyze_text(
                company_name, combined_prompt_content, PROMPT_COMBINED_ANALYSIS, together_client)
    else:  # Fallback if individual_analyses_list is empty, though already handled at function start
        overall_summary_result = "Overall No Churn Risk Indicated\nReason: Not Applicable - No relevant churn indicators found.\nSummary: No relevant news articles found for overall analysis."

    return {"individual_analyses": individual_analyses_list, "overall_summary": overall_summary_result}


def get_risk_level_from_text(text_summary: str) -> str:
    """
    Extracts risk level from a summary string by looking for the explicit risk phrases
    at the beginning of the text, robustly handling "Overall" prefix.
    """
    lines = text_summary.strip().split('\n')
    if not lines:
        return "Unknown Risk"

    first_line_lower = lines[0].lower()

    # Check for specific risk levels
    if "high risk" in first_line_lower:  # Catches "High Risk" and "Overall High Risk"
        return "High Risk"
    elif "medium risk" in first_line_lower:  # Catches "Medium Risk" and "Overall Medium Risk"
        return "Medium Risk"
    elif "low risk" in first_line_lower:  # Catches "Low Risk" and "Overall Low Risk"
        return "Low Risk"
    # Catches "No Churn Risk Indicated" and "Overall No Churn Risk Indicated"
    elif "no churn risk indicated" in first_line_lower:
        return "No Churn Risk Indicated"

    return "Unknown Risk"


def extract_reason_from_ai_summary(ai_summary: str) -> str:
    """
    Extracts the 'Reason(s) for Risk' line from the AI's summary,
    expecting it to start with 'Reason:'.
    """
    lines = ai_summary.strip().split('\n')
    for line in lines:
        if line.lower().startswith("reason:"):
            # Return everything after "Reason: "
            return line[len("reason:"):].strip()
    return "Reason not explicitly stated by AI."


def display_summary_with_color(company_name, summary_text):
    """Displays the summary with color coding based on risk level."""
    risk_level = get_risk_level_from_text(summary_text)

    st.markdown(f"### Summary for {company_name}")

    if "High Risk" in risk_level:
        st.error(summary_text)
    elif "Medium Risk" in risk_level:
        st.warning(summary_text)
    elif "Low Risk" in risk_level:
        st.info(summary_text)
    else:  # For "No Churn Risk Indicated" and "Unknown Risk"
        st.success(summary_text)


def run_analysis(company_data_df, days_to_search):
    """Main function to orchestrate the news fetching and analysis for multiple companies."""
    results = {}
    today = datetime.today()
    from_date = today - timedelta(days=days_to_search)
    max_articles_per_query = 10

    churn_keywords = {
        "Corporate Restructuring": ["merger", "acquisition", "investment", "joint venture", "IPO", "restructuring", "realignment", "rebranding", "subsidiary", "consolidation"],
        "Business Discontinuity": ["shutdown", "closed", "bankruptcy", "insolvency", "pivot", "market exit"],
        "Strategic Policy Changes": ["benefits withdrawn", "benefits discontinued", "centralization", "new CEO", "cost cutting", "budget cuts", "strategy shift"],
        "Financial Constraints": ["payroll issue", "financial loss", "cost pressure", "cash flow", "budget reallocation"],
        "Employment Structure Changes": ["employee transfer", "contractual workforce", "remote work", "layoffs", "furloughs", "downsizing"],
        "Regulatory & Compliance": ["tax policy", "labor law", "income tax", "GST change", "budget amendment", "social security"],
        "Competitive Market Dynamics": ["switched vendor", "new platform", "competitor", "pricing", "market share", "disruption", "value proposition"],
        "Technological Transitions": ["digital transformation", "HRMS integration", "API", "analytics", "mobile app", "platform upgrade"],
        "Service Delivery Issues": ["onboarding delay", "tech issues", "merchant issue", "support problem", "delivery delay", "reimbursement issue"],
        "Employee Engagement": ["low adoption", "user experience", "employee feedback", "generation gap", "hybrid work", "usage drop"]
    }

    allowed_domains = [
        "livemint.com", "economictimes.indiatimes.com", "business-standard.com",
        "thehindubusinessline.com", "financialexpress.com", "ndtvprofit.com",
        "zeebiz.com", "moneycontrol.com", "bloombergquint.com",
        "cnbctv18.com", "businesstoday.in", "indianexpress.com",
        "thehindu.com", "reuters.com", "businesstraveller.com",
        "sify.com", "telegraphindia.com", "outlookindia.com",
        "firstpost.com", "pulse.zerodha.com", "ndtvprofit.com",
        "ddnews.gov.in", "newsonair.gov.in", "pib.gov.in",
        "niti.gov.in", "rbi.org.in", "sebi.gov.in",
        "dpiit.gov.in", "investindia.gov.in", "indiabriefing.com",
        "Taxscan.in", "bwbusinessworld.com", "inc42.com",
        "yourstory.com", "vccircle.com", "entrackr.com",
        "the-ken.com", "linkedin.com", "mca.gov.in",
        "zaubacorp.com", "tofler.in"
    ]
    processed_allowed_domains = [domain.replace(
        "www.", "") for domain in allowed_domains]

    st.sidebar.subheader("Analysis Parameters")
    st.sidebar.info(
        f"Analyzing news from: **{from_date.strftime('%Y-%m-%d')}** to **{today.strftime('%Y-%m-%d')}** ({days_to_search} days)")
    st.sidebar.info(f"Max Articles per Query: **{max_articles_per_query}**")
    st.sidebar.info(
        f"Filtered by {len(processed_allowed_domains)} specified business news domains.")

    for index, row in company_data_df.iterrows():
        company_name = row["CompanyName"]
        # Ensure it's a string and strip whitespace
        provided_churn_reason = str(row["ChurnReason"]).strip()

        queries = [company_name] + [f"{company_name} {keyword}" for category_keywords in churn_keywords.values()
                                    for keyword in category_keywords]
        company_analysis = analyze_news(
            company_name, from_date, today, max_articles_per_query, queries, processed_allowed_domains
        )

        # Extract AI's primary reason for comparison
        ai_overall_summary = company_analysis.get(
            "overall_summary", "Overall No Churn Risk Indicated\nReason: Not Applicable - No analysis available.")
        ai_extracted_reason = extract_reason_from_ai_summary(
            ai_overall_summary)

        # Compare reasons using LLM
        comparison_result = get_llm_reason_comparison_result(
            ai_extracted_reason, provided_churn_reason, together_client)

        results[company_name] = {
            "analysis": company_analysis,
            "provided_churn_reason": provided_churn_reason,
            "ai_extracted_reason": ai_extracted_reason,
            "comparison_result": comparison_result
        }
    return results


# --- Streamlit App Layout ---
st.set_page_config(page_title="Company Churn Risk Analyzer", layout="wide")
st.title("💡 Company Churn Risk Analysis (India Focus)")
st.markdown("""
This application helps identify potential client churn risks for an employee benefits company in India by analyzing recent news articles.
Upload an **Excel file** containing 'CompanyName' and 'ChurnReason' columns, and the app will fetch relevant news, summarize it, and compare the AI's predicted reason with your provided reason.
""")

# User input for number of days
days_to_search = st.slider(
    "Select the number of days for news search (back from today):",
    min_value=1,
    max_value=365,
    value=90,  # Default to 90 days
    step=1,
    help="This determines how far back in time the news articles will be fetched."
)

# File uploader widget for Excel
uploaded_file = st.file_uploader(
    "Upload your Excel file with 'CompanyName' and 'ChurnReason' columns", type=["xlsx"])

company_data_df = pd.DataFrame()
if uploaded_file is not None:
    try:
        company_df = pd.read_excel(uploaded_file)
        if "CompanyName" in company_df.columns and "ChurnReason" in company_df.columns:
            company_data_df = company_df.dropna(
                subset=["CompanyName"]).reset_index(drop=True)
            if not company_data_df.empty:
                st.success(
                    f"Successfully loaded **{len(company_data_df)}** companies from **'{uploaded_file.name}'**.")
                st.info("You can now click 'Start Analysis' to begin.")
            else:
                st.warning(
                    "The 'CompanyName' column is empty after loading. Please check your Excel file.")
        else:
            st.error(
                "Error: The uploaded Excel must contain **'CompanyName'** and **'ChurnReason'** columns.")
    except Exception as e:
        st.error(
            f"Error reading Excel file: {e}. Please ensure it's a valid XLSX file with the correct column names.")
else:
    st.info("Please upload an Excel file with 'CompanyName' and 'ChurnReason' columns to proceed.")


if st.button("🚀 Start Analysis"):
    if company_data_df.empty:
        st.warning(
            "No company data available for analysis. Please upload an Excel file.")
    else:
        with st.spinner("Crunching numbers and fetching news... This might take a while for each company."):
            analysis_results = run_analysis(
                company_data_df, days_to_search)

        st.success("🎉 Analysis Complete!")
        st.markdown("---")

        # Display the results
        comparison_counts = {"Match": 0, "Partial Match": 0,
                             "No Match": 0, "No Match (LLM Error)": 0}

        for company, data in analysis_results.items():
            analysis = data["analysis"]
            provided_churn_reason = data["provided_churn_reason"]
            ai_extracted_reason = data["ai_extracted_reason"]
            comparison_result = data["comparison_result"]

            if comparison_result in comparison_counts:
                comparison_counts[comparison_result] += 1
            else:
                # Catch any unexpected comparison results from LLM
                comparison_counts["No Match"] += 1

            st.markdown(f"## :office: {company}")

            # Display Overall Churn Risk Summary with color and new heading
            display_summary_with_color(company, analysis.get(
                "overall_summary", "Overall No Churn Risk Indicated\nReason: Not Applicable - No analysis available."))

            st.markdown(
                f"**Provided Churn Reason (from Excel):** `{provided_churn_reason}`")
            st.markdown(f"**AI Extracted Reason:** `{ai_extracted_reason}`")
            st.markdown(f"**Reason Comparison:** **`{comparison_result}`**")

            st.markdown("### Individual Article Analyses")
            if analysis.get("individual_analyses"):
                for i, article_analysis in enumerate(analysis["individual_analyses"]):
                    st.markdown(
                        f"#### :newspaper: {article_analysis['title']}")
                    st.markdown(f"**URL:** [Link]({article_analysis['url']})")
                    article_analysis_text = article_analysis['analysis']

                    # Color-code individual analysis using the refined get_risk_level_from_text
                    individual_risk_level = get_risk_level_from_text(
                        article_analysis_text)
                    if "High Risk" in individual_risk_level:
                        st.error(f"**Analysis:** {article_analysis_text}")
                    elif "Medium Risk" in individual_risk_level:
                        st.warning(f"**Analysis:** {article_analysis_text}")
                    elif "Low Risk" in individual_risk_level:
                        st.info(f"**Analysis:** {article_analysis_text}")
                    else:
                        st.write(f"**Analysis:** {article_analysis_text}")
                    st.markdown("---")
            else:
                st.info("No individual articles found for detailed analysis.")
            st.markdown("---")  # Separator between companies

        # Display Percentage of Results in UI
        st.subheader("Comparison Result Summary")
        total_companies = len(analysis_results)
        if total_companies > 0:
            for result_type, count in comparison_counts.items():
                percentage = (count / total_companies) * 100
                st.markdown(
                    f"- **{result_type}**: {count} companies ({percentage:.2f}%)")
            st.markdown("---")
        else:
            st.info("No companies analyzed to show comparison summary.")
            st.markdown("---")

        # Export to Excel (with a new sheet for percentages)
        data_for_df = []
        for company, data in analysis_results.items():
            analysis = data["analysis"]
            overall_summary = analysis.get(
                "overall_summary", "Overall No Churn Risk Indicated\nReason: Not Applicable - No analysis available.")
            overall_risk_level = get_risk_level_from_text(overall_summary)

            company_data = {
                "Company": company,
                "Overall Risk Level": overall_risk_level,
                "Overall Summary": overall_summary,
                "Provided Churn Reason (from Excel)": data["provided_churn_reason"],
                "AI Extracted Reason": data["ai_extracted_reason"],
                "Comparison Result": data["comparison_result"]
            }
            for i, article_analysis in enumerate(analysis.get("individual_analyses", [])):
                article_risk_level = get_risk_level_from_text(
                    article_analysis["analysis"])
                company_data[f"Article {i+1} Title"] = article_analysis["title"]
                company_data[f"Article {i+1} URL"] = article_analysis["url"]
                company_data[f"Article {i+1} Risk Level"] = article_risk_level
                company_data[f"Article {i+1} Analysis"] = article_analysis["analysis"]
            data_for_df.append(company_data)

        if data_for_df:
            df_results = pd.DataFrame(data_for_df)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            excel_file_name = f"churn_analysis_results_{timestamp}.xlsx"

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                # Write main results to 'Detailed Analysis' sheet
                df_results.to_excel(
                    writer, sheet_name='Detailed Analysis', index=False)

                # Prepare and write summary percentages to 'Comparison Summary' sheet
                summary_data = []
                total_companies = len(analysis_results)
                if total_companies > 0:
                    for result_type, count in comparison_counts.items():
                        percentage = (count / total_companies) * 100
                        summary_data.append(
                            {"Result Type": result_type, "Count": count, "Percentage": f"{percentage:.2f}%"})
                else:
                    summary_data.append(
                        {"Result Type": "No Data", "Count": 0, "Percentage": "0.00%"})

                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(
                    writer, sheet_name='Comparison Summary', index=False)

            excel_buffer.seek(0)

            st.download_button(
                label="Download All Results as Excel 📊",
                data=excel_buffer,
                file_name=excel_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the comprehensive analysis results."
            )
            st.success("Results are ready for download!")
        else:
            st.warning(
                "No data to export to Excel, as no analysis was performed.")

st.sidebar.markdown("---")
st.sidebar.header("About")
st.sidebar.info(
    "This app leverages Together AI's Llama-3.3-70B-Instruct-Turbo-Free model and Google News for churn risk analysis.")
