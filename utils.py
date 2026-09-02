import os
import json
import re
import threading
import PyPDF2
import docx
import requests
import base64
import datetime
from io import BytesIO
from openai import OpenAI
import gemini_client
from fpdf import FPDF
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import numpy as np
from spellchecker import SpellChecker

_spell_checker = None

def _get_spell_checker():
    global _spell_checker
    if _spell_checker is None:
        _spell_checker = SpellChecker()
    return _spell_checker


def _is_valid_english_word(word):
    sc = _get_spell_checker()
    w = word.lower().strip().strip(".,;:!?\"'()[]{}…–—")
    if not w or not w.isalpha():
        return True
    if w in _get_common_words():
        return True
    known = sc.known([w])
    return w in known


def _validate_spelling_finding(finding):
    orig = finding.get("original_text", "").strip()
    suggestion = finding.get("suggestion", "").strip()

    if _is_american_british_pair(orig, suggestion):
        return True

    orig_joined = re.sub(r'\s+', '', orig).lower()
    sugg_joined = re.sub(r'\s+', '', suggestion).lower()

    if orig_joined == sugg_joined:
        return False

    if orig_joined.isalpha() and _is_valid_english_word(orig_joined):
        if sugg_joined.isalpha() and _is_valid_english_word(sugg_joined):
            return False
        return False

    words = orig.lower().split()
    if len(words) > 1 and all(_is_valid_english_word(w) for w in words if w.isalpha()):
        sugg_words = suggestion.lower().split()
        if len(sugg_words) >= 1:
            if all(_is_valid_english_word(w) for w in sugg_words if w.isalpha()):
                return False

    return True


_AMERICAN_TO_BRITISH = {
    "organization": "organisation", "organizations": "organisations",
    "organize": "organise", "organized": "organised", "organizing": "organising",
    "center": "centre", "centers": "centres",
    "analyze": "analyse", "analyzed": "analysed", "analyzing": "analysing",
    "color": "colour", "colors": "colours", "colored": "coloured",
    "favor": "favour", "favored": "favoured", "favorable": "favourable",
    "honor": "honour", "honored": "honoured",
    "labor": "labour",
    "program": "programme", "programs": "programmes",
    "defense": "defence",
    "offense": "offence",
    "license": "licence",
    "practice": "practise",
    "realize": "realise", "realized": "realised", "realizing": "realising",
    "recognize": "recognise", "recognized": "recognised", "recognizing": "recognising",
    "specialize": "specialise", "specialized": "specialised",
    "utilize": "utilise", "utilized": "utilised", "utilizing": "utilising",
    "standardize": "standardise", "standardized": "standardised",
    "modernize": "modernise", "modernized": "modernised",
    "privatize": "privatise", "privatized": "privatised",
    "liberalize": "liberalise", "liberalized": "liberalised",
    "decentralize": "decentralise", "decentralized": "decentralised",
    "centralize": "centralise", "centralized": "centralised",
    "digitize": "digitise", "digitized": "digitised",
    "customize": "customise", "customized": "customised",
    "optimize": "optimise", "optimized": "optimised",
    "maximize": "maximise", "maximized": "maximised",
    "minimize": "minimise", "minimized": "minimised",
    "prioritize": "prioritise", "prioritized": "prioritised",
    "behavior": "behaviour", "behaviors": "behaviours",
    "catalog": "catalogue",
    "dialog": "dialogue",
    "judgment": "judgement",
    "traveling": "travelling",
    "counseling": "counselling",
    "modeling": "modelling",
    "enrollment": "enrolment",
    "fulfillment": "fulfilment",
    "installment": "instalment",
    "neighborhood": "neighbourhood",
    "acknowledgment": "acknowledgement",
}

_BRITISH_TO_AMERICAN = {v: k for k, v in _AMERICAN_TO_BRITISH.items()}


def _is_american_british_pair(orig, suggestion):
    o = orig.lower().strip()
    s = suggestion.lower().strip()
    if o in _AMERICAN_TO_BRITISH and _AMERICAN_TO_BRITISH[o] == s:
        return True
    if o in _BRITISH_TO_AMERICAN and _BRITISH_TO_AMERICAN[o] == s:
        return True
    if o.endswith("ize") and s == o[:-3] + "ise":
        return True
    if o.endswith("ization") and s == o[:-7] + "isation":
        return True
    if o.endswith("izing") and s == o[:-5] + "ising":
        return True
    if o.endswith("ized") and s == o[:-4] + "ised":
        return True
    if o.endswith("or") and s == o[:-2] + "our" and len(o) >= 6:
        if o in _AMERICAN_TO_BRITISH:
            return True
    if o.endswith("ense") and s == o[:-4] + "ence":
        return True
    if o.endswith("og") and s == o + "ue":
        return True
    return False

# OpenAI client (default provider), created on first use.
# When USE_GEMINI is enabled, call_openai_api() routes to Gemini instead and
# this client is never built, so OPENAI_API_KEY is not required in that mode.
# The client is created lazily because recent openai SDKs raise on a missing
# API key at construction time, which would break import in Gemini-only
# deployments.
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
openai_client = None
_openai_client_lock = threading.Lock()


def get_openai_client():
    """Return the process-wide OpenAI client, building it on first call."""
    global openai_client
    if openai_client is not None:
        return openai_client
    with _openai_client_lock:
        if openai_client is None:
            # `or None` lets the SDK fall back to its own env-var lookup.
            openai_client = OpenAI(api_key=OPENAI_API_KEY or None)
        return openai_client

# Prefixes of the error strings returned by call_openai_api() on failure.
API_ERROR_PREFIXES = ("Error calling OpenAI API:", "Error calling Gemini API:")


def is_api_error_message(value):
    """True when a call_openai_api() result is a provider error string."""
    return isinstance(value, str) and value.startswith(API_ERROR_PREFIXES)


def describe_ai_provider():
    """One-line description of the AI provider this process will actually use.

    Contains no secrets — keys are reported as set/missing, never echoed.
    """
    if gemini_client.use_gemini():
        return gemini_client.describe_provider()
    return "OpenAI | model=gpt-4o (default) | OPENAI_API_KEY=%s" % (
        "set" if OPENAI_API_KEY else "MISSING",
    )


# Logged once per process so deployment logs state unambiguously which
# provider is serving traffic.
print("[AI provider] " + describe_ai_provider())

_COMMON_WORDS = None

def _get_common_words():
    global _COMMON_WORDS
    if _COMMON_WORDS is not None:
        return _COMMON_WORDS
    _COMMON_WORDS = {
        "the","of","and","to","in","a","is","that","for","it","was","on","are","as","with",
        "his","they","be","at","one","have","this","from","or","had","by","not","but","what",
        "all","were","we","when","your","can","said","there","each","which","she","do","how",
        "their","if","will","up","other","about","out","many","then","them","these","so",
        "some","her","would","make","like","has","him","into","time","very","an","been","no",
        "its","who","did","get","may","than","now","people","my","over","such","our","also",
        "new","after","us","use","two","way","well","state","more","most","only","any","work",
        "first","just","where","case","back","much","go","good","give","could","study","own",
        "take","come","made","find","here","thing","long","still","between","right","too","old",
        "tell","year","public","before","should","because","through","same","part","last","while",
        "under","never","great","know","must","those","need","house","government","service",
        "shall","country","hand","school","national","world","high","every","system","area",
        "i","me","he","she","it","we","you","am","are","is","was","were","be","been","being",
        "have","has","had","do","does","did","will","would","shall","should","may","might",
        "can","could","must","ought","need","dare","used","going","being","having","doing",
        "able","go","see","say","get","put","set","run","let","try","ask","own","end","big",
        "far","old","low","few","off","day","man","way","eye","boy","did","got","cut","lot",
        "bit","age","top","war","god","per","act","law","bed","key","tax","eat","car","air",
        "job","art","cup","oil","sat","dog","led","red","hot","bad","six","sea","add","won",
        "nor","due","pay","met","arm","nor","yet","hit","sit","buy","die","lay","fit","ten",
        "map","raw","aim","aid","gap","era","via","tip","gas",
        "in","on","at","to","by","up","an","or","no","so","if","as","do","my","he","we",
        "built","structure","plan","down","left","full","life","door","real","show",
        "keep","turn","move","point","four","five","open","city","land","form","line",
        "body","small","large","number","change","place","found","head","page","kind",
        "food","water","stand","call","look","think","came","again","help","play","never",
        "three","seven","eight","nine","once","both","next","face","left","town","even",
        "free","sure","half","hold","near","side","took","above","along","class","close",
        "name","best","upon","past","done","gone","less","single","second","third",
        "start","force","power","total","clear","round","order","given","data","level",
        "local","model","human","field","board","court","chief","legal","urban","rural",
        "civil","cost","fund","growth","rate","cent","plan","phase","step","role",
        "rule","term","type","unit","fact","view","rest","sent","rise","note","lead",
        "base","wide","deep","care","mark","sure","size","test","post","bank","rate",
        "risk","sale","loss","link","firm","loan","deal","bond","debt","flow","mode",
        "core","code","file","tool","sign","rank","poll","seat","vote","bill","item",
        "told","felt","left","kept","gave","sent","went","held","took","knew","seen",
    }
    _COMMON_WORDS |= {
        "governance","administration","implementation","organisation","organization",
        "infrastructure","development","management","government","department",
        "initiative","institutional","accountability","transparency","competency",
        "competencies","efficiency","effectiveness","sustainability","stakeholder",
        "stakeholders","collaboration","innovation","technology","digitization",
        "digitisation","digitalisation","digitalization","bureaucratic","bureaucracy",
        "decentralisation","decentralization","empowerment","transformation",
        "modernisation","modernization","rehabilitation","procurement","expenditure",
        "jurisdiction","administrative","commissioner","municipality","municipal",
        "panchayat","surveillance","enforcement","compliance","regulation",
        "regulatory","legislature","legislative","parliamentary","constitutional",
        "resolution","intervention","mechanism","framework","programme","program",
        "evaluation","assessment","achievement","performance","improvement",
        "enhancement","recommendation","documentation","communication","information",
        "distribution","contribution","participation","representation","recognition",
        "integration","preparation","investigation","consideration","notification",
        "application","examination","determination","discrimination","authorization",
        "certification","classification","specification","identification","demonstration",
        "accommodation","approximation","characterisation","characterization",
        "privatisation","privatization","liberalisation","liberalization",
        "commercialisation","commercialization","industrialisation","industrialization",
        "rationalisation","rationalization","standardisation","standardization",
        "centralisation","centralization","marginalisation","marginalization",
        "operationalisation","operationalization","professionalisation",
        "agriculture","agricultural","education","educational","environment",
        "environmental","population","technological","geographical","geographical",
        "significant","significance","opportunity","opportunities","responsibility",
        "responsibilities","understanding","approximately","comprehensive",
        "establishment","particularly","international","independently","unfortunately",
        "subsequently","simultaneously","infrastructure","characteristics","consequently",
        "transportation","communication","communications","pharmaceutical","manufacturing",
        "predominantly","circumstances","transformation","transformative","supplementary",
        "approximately","extraordinary","unprecedented","disadvantaged","disproportionate",
        "acknowledgement","acknowledgment","neighbourhood","neighborhood",
        "programme","programmes","analyse","analysed","analysing","organisation",
        "organisations","organised","organising","recognised","recognising",
        "behaviour","behaviours","favour","favoured","labour","colour","honour",
        "centre","centres","defence","licence","offence","practice","practise",
        "catalogue","dialogue","cheque","grey","judgement","manoeuvre","plough",
        "programme","sceptical","specialise","specialised","travelling","counselling",
        "modelling","cancellation","jewellery","skilful","wilful","enrolment",
        "fulfilment","instalment","focussed","focussing",
        "accountability","egovernance","e-governance","mgovernance","m-governance",
        "karmayogi","igot","swachh","bharat","pradhan","mantri","niti","aayog",
        "ayushman","ujjwala","panchayati","seva","yojana","andolan","swaraj",
        "adhikari","sabha","samiti","zila","parishad","tehsil","taluka","nagar",
        "nigam","palika","vidhan","rajya","lok","vibhag","mandal","kshetra",
        "mission","national","district","collector","magistrate","commissioner",
        "secretary","directorate","secretariat","ministry","division","bureau",
        "tribunal","appellate","ombudsman","vigilance","comptroller","auditor",
        "exchequer","gazette","notification","ordinance","statute","amendment",
    }
    return _COMMON_WORDS


def _is_likely_word(token, common_words):
    t = token.lower()
    if len(t) <= 1:
        return True
    if t in common_words:
        return True
    if t.endswith("'s") and t[:-2] in common_words:
        return True
    for suffix in ("ed","ing","tion","sion","ment","ness","ity","ise","ize","ised","ized",
                    "ising","izing","able","ible","ful","less","ous","ive","al","ly","er","est","s","es"):
        if not t.endswith(suffix):
            continue
        stem = t[:-len(suffix)]
        if len(stem) >= 3 and stem in common_words:
            return True
    return False


def repair_broken_words(text):
    common_words = _get_common_words()

    words_in_doc = set()
    for m in re.finditer(r"[A-Za-z]{5,}", text):
        words_in_doc.add(m.group().lower())

    def _joined_is_valid(joined):
        jl = joined.lower()
        return (jl in common_words
                or jl in words_in_doc
                or _is_likely_word(joined, common_words)
                or _is_valid_english_word(jl))

    def fix_line(line):
        tokens = re.split(r'(\s+)', line)
        result = []
        i = 0
        while i < len(tokens):
            tok = tokens[i]
            if not re.match(r'^[A-Za-z]+$', tok) or len(tok) == 0:
                result.append(tok)
                i += 1
                continue

            _COMMON_SHORT_WORDS = {
                "a","i","an","am","as","at","be","by","do","go","he","if","in",
                "is","it","me","my","no","of","oh","ok","on","or","so","to","up",
                "us","we","the","and","are","but","can","did","for","get","got",
                "had","has","her","him","his","how","its","let","may","new","nor",
                "not","now","odd","off","old","one","our","out","own","per","put",
                "ran","run","saw","say","set","she","sit","six","ten","the","too",
                "two","use","was","way","who","why","won","yet","you","log","move",
                "check",
            }

            best_joined = tok
            best_end = i + 1
            j = i
            accumulated = tok
            while j + 2 < len(tokens):
                if (tokens[j+1].strip() == ''
                    and len(tokens[j+1]) <= 2
                    and re.match(r'^[A-Za-z]+$', tokens[j+2])):
                    candidate = accumulated + tokens[j+2]
                    j += 2
                    accumulated = candidate
                    if _joined_is_valid(candidate):
                        frags = []
                        k = i
                        while k <= j:
                            if re.match(r'^[A-Za-z]+$', tokens[k]):
                                frags.append(tokens[k])
                            k += 1
                        all_frags_words = all(_is_likely_word(f, common_words) for f in frags)
                        all_common_short = all(f.lower() in _COMMON_SHORT_WORDS for f in frags)
                        if all_common_short:
                            pass
                        elif not all_frags_words or len(frags) >= 3 or any(len(f) <= 2 and f.lower() not in _COMMON_SHORT_WORDS for f in frags):
                            best_joined = candidate
                            best_end = j + 1
                else:
                    break

            if best_end > i + 1:
                result.append(best_joined)
                i = best_end
                continue

            if (i + 2 < len(tokens)
                and tokens[i+1].strip() == ''
                and len(tokens[i+1]) <= 2
                and re.match(r'^[A-Za-z]+$', tokens[i+2])):

                frag1 = tok
                frag2 = tokens[i+2]
                joined = frag1 + frag2

                f1_is_word = _is_likely_word(frag1, common_words) or _is_valid_english_word(frag1)
                f2_is_word = _is_likely_word(frag2, common_words) or _is_valid_english_word(frag2)
                joined_is_word = _joined_is_valid(joined)

                if f1_is_word and f2_is_word:
                    result.append(tok)
                    i += 1
                    continue

                should_join = False
                if joined_is_word and not f1_is_word and not f2_is_word:
                    should_join = True
                if not f1_is_word and not f2_is_word and len(frag1) <= 4 and joined_is_word:
                    should_join = True
                if not f1_is_word and len(frag1) <= 2 and frag1.lower() not in _COMMON_SHORT_WORDS and joined_is_word:
                    should_join = True
                if len(frag1) == 1 and frag1.lower() not in _COMMON_SHORT_WORDS and joined_is_word:
                    should_join = True

                if should_join:
                    result.append(joined)
                    i += 3
                    continue

            result.append(tok)
            i += 1
        return ''.join(result)

    lines = text.split('\n')
    repaired = [fix_line(line) for line in lines]

    text = '\n'.join(repaired)

    text = text.replace('ﬃ', 'ffi')
    text = text.replace('ﬄ', 'ffl')
    text = text.replace('ﬁ', 'fi')
    text = text.replace('ﬂ', 'fl')
    text = text.replace('ﬀ', 'ff')

    return text


def normalize_extracted_text(text):
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'(\w) +- +(\w)', r'\1-\2', text)
    text = re.sub(r'(\w) +-(\w)', r'\1-\2', text)
    text = re.sub(r'(\w)- +(\w)', r'\1-\2', text)
    text = re.sub(r' +([,.:;!?\)])', r'\1', text)
    text = re.sub(r'(\() +', r'\1', text)
    text = re.sub(r'" +', '"', text)
    text = re.sub(r' +"', ' "', text)
    text = re.sub(r'\u2019 +', '\u2019', text)
    text = re.sub(r'\u201c +', '\u201c', text)
    text = re.sub(r' +\u201d', '\u201d', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    lines = text.split('\n')
    cleaned_lines = [line.strip() for line in lines]
    text = '\n'.join(cleaned_lines)
    text = repair_broken_words(text)
    return text


def extract_text_from_pdf(pdf_file):
    """
    Extract text content from a PDF file.
    
    Args:
        pdf_file: The uploaded PDF file object
        
    Returns:
        str: Extracted text from the PDF
    """
    text = ""
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n\n"
    except Exception as e:
        raise Exception(f"Error extracting text from PDF: {str(e)}")
    
    return normalize_extracted_text(text)

def extract_text_from_docx(docx_file):
    """
    Extract text content from a DOCX file.
    
    Args:
        docx_file: The uploaded DOCX file object
        
    Returns:
        str: Extracted text from the DOCX
    """
    text = ""
    try:
        # Convert file to BytesIO object
        bytes_content = BytesIO(docx_file.getvalue())
        
        # Open using python-docx
        doc = docx.Document(bytes_content)
        
        # Extract text from paragraphs
        for para in doc.paragraphs:
            text += para.text + "\n"
        
        # Extract text from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text += cell.text + " "
                text += "\n"
        
    except Exception as e:
        raise Exception(f"Error extracting text from DOCX: {str(e)}")
    
    return text

def call_openai_api(prompt, model="gpt-4o", response_format=None, temperature=0.5, seed=None):
    """
    Call the configured AI provider with the given prompt.

    This is the single entry point for every AI call in the application. The
    provider is chosen by the USE_GEMINI environment variable:
      * unset/false -> OpenAI chat-completions (default, unchanged behaviour)
      * true        -> Google Gemini (Vertex AI or Gemini Developer API)
    Both paths return the same shapes, so callers need no changes.

    Args:
        prompt (str): The prompt to send to the API
        model (str): Model to use. In Gemini mode an OpenAI model id such as
            "gpt-4o" is mapped to GENAI_MODEL_NAME; an explicit Gemini model
            id is honoured as given.
        response_format (str, optional): Format for response (e.g., "json_object")
        temperature (float): Temperature for response generation (default 0.5)
        seed (int, optional): Fixed seed for deterministic output

    Returns:
        The parsed dict when response_format == "json_object", otherwise the
        response text. On failure, an "Error calling <provider> API: ..."
        string (see is_api_error_message()).
    """
    use_gemini = gemini_client.use_gemini()
    try:
        if use_gemini:
            # Gemini equivalents of the OpenAI arguments (JSON mode via
            # response_mime_type, max_output_tokens, markdown-fence stripping)
            # are handled inside gemini_client.generate_text().
            content = gemini_client.generate_text(
                prompt,
                model=model,
                response_format=response_format,
                temperature=temperature,
                seed=seed,
            )
        else:
            messages = [{"role": "user", "content": prompt}]

            kwargs = {}
            if response_format == "json_object":
                kwargs["response_format"] = {"type": "json_object"}
            if seed is not None:
                kwargs["seed"] = seed

            response = get_openai_client().chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=2000,
                **kwargs
            )

            # Process the response
            content = response.choices[0].message.content
        
        # If JSON format was requested, parse the result
        if response_format == "json_object":
            try:
                result = json.loads(content)
                # Gemini can return a top-level array where OpenAI's
                # json_object mode always returns an object; treat that like a
                # parse failure so the fallback shape below is used.
                if not isinstance(result, dict):
                    raise ValueError("Expected a JSON object at the top level")
                
                # Ensure all expected fields exist and have the correct types
                if "score" not in result:
                    result["score"] = 0  # Default to 0 for case study scoring
                elif not isinstance(result["score"], (int, float)):
                    # Convert to number if possible, otherwise use default
                    try:
                        result["score"] = float(result["score"])
                    except ValueError:
                        result["score"] = 0
                
                # Ensure reasoning is a string
                if "reasoning" not in result:
                    result["reasoning"] = "No reasoning provided"
                elif isinstance(result["reasoning"], list):
                    result["reasoning"] = ". ".join([str(item) for item in result["reasoning"]])
                
                # Ensure document_reference is a string
                if "document_reference" not in result:
                    result["document_reference"] = "No document reference provided"
                elif isinstance(result["document_reference"], list):
                    result["document_reference"] = ". ".join([str(item) for item in result["document_reference"]])
                
                return result
            except ValueError:
                # ValueError covers json.JSONDecodeError and the non-object
                # guard above.
                return {
                    "score": 0,
                    "reasoning": "Failed to parse response from AI model",
                    "document_reference": "Error in analysis"
                }
        
        return content
    
    except Exception as e:
        provider = "Gemini" if use_gemini else "OpenAI"
        return f"Error calling {provider} API: {str(e)}"


def analyze_writing_quality_chunked(full_text, exclusion_instructions="", chunk_size=3000, overlap=500, max_workers=5):
    from concurrent.futures import ThreadPoolExecutor, as_completed

    text_length = len(full_text)
    if text_length == 0:
        return []

    chunks = []
    start = 0
    while start < text_length:
        end = min(start + chunk_size, text_length)
        chunks.append((start, full_text[start:end]))
        if end >= text_length:
            break
        start += chunk_size - overlap

    total_chunks = len(chunks)

    def process_chunk(chunk_idx, chunk_text):
        prompt = f"""You are an expert English language editor and proofreader specialising in academic and government case studies written in British English.

Analyse the following text chunk (chunk {chunk_idx + 1} of {total_chunks}) for writing issues. Be HIGHLY ACCURATE — only flag genuine errors, not stylistic preferences. Do NOT flag proper nouns, acronyms, abbreviations, or domain-specific terminology.

CHECK FOR THESE SPECIFIC ISSUE TYPES:

1. **Spelling** — Identify words that are misspelt. Flag American English spellings that should be British English (e.g., "organization" → "organisation", "center" → "centre", "analyze" → "analyse", "color" → "colour", "program" → "programme", "defense" → "defence", "favor" → "favour", "realize" → "realise").

2. **Grammar** — Check subject-verb agreement errors, incorrect article usage, missing or extra prepositions, incorrect pronoun references, and misused words.

3. **Tense Consistency** — Case studies should primarily use past tense. Flag any inconsistent shifts to present tense within narrative sections (but allow present tense in general statements, conclusions, or when describing ongoing situations).

4. **Redundancy** — Flag redundant phrases like "in order to" (use "to"), "due to the fact that" (use "because"), "at this point in time" (use "now"), "each and every" (use "each" or "every"), "free gift" (use "gift"), repeated ideas in adjacent sentences.

5. **Sentence Structure** — Flag overly complex sentences (more than 40 words with multiple clauses), sentences in passive voice that would be clearer in active voice, run-on sentences, and suggest simpler alternatives.

{exclusion_instructions}

CRITICAL — PDF EXTRACTION ARTIFACT RULES (DO NOT VIOLATE):
This text was extracted from a PDF document. The extraction process sometimes introduces artefacts that are NOT present in the original document. You MUST NOT flag any of the following as errors:
- Words that appear broken with extra spaces (e.g., "gover nance", "admini stration", "imple mentation", "organi sation") — these are extraction artefacts where the original PDF has the word correctly spelled
- Spaces around hyphens or dashes (e.g., "e -governance" instead of "e-governance")
- Ligature splitting (e.g., "ﬁ" becoming "fi", "ﬂ" becoming "fl")
- Irregular whitespace, missing spaces between words, or extra spaces within words
- Any "spelling error" where the suggestion is simply the same letters with different spacing
- Any finding where removing all spaces from the original and the suggestion produces the same text
If a word looks misspelled but could be a PDF extraction artefact (broken word), DO NOT flag it. When in doubt, DO NOT flag it.

=== TEXT TO ANALYSE ===
{chunk_text}

Return your findings as a JSON object with this EXACT structure:
{{
    "findings": [
        {{
            "type": "Spelling|Grammar|Tense|Redundancy|Structure",
            "original_text": "the exact problematic text from the document (keep short, max 15 words)",
            "suggestion": "the corrected or improved version",
            "severity": "High|Medium|Low",
            "context": "brief explanation of the issue (one sentence)"
        }}
    ]
}}

Rules:
- Return ONLY genuine issues — do NOT fabricate or invent findings
- If there are no issues in this chunk, return {{"findings": []}}
- Keep original_text short (the specific problematic phrase only, max 15 words)
- severity: High = clear error (spelling, grammar), Medium = tense inconsistency or significant redundancy, Low = style improvement or minor redundancy
- Do NOT flag proper nouns, place names, organisation names, or technical terms
- Do NOT flag direct quotes from other sources
- Be conservative — when in doubt, do NOT flag it"""

        result = call_openai_api(prompt, response_format="json_object", temperature=0.1, seed=42)
        findings = []
        if isinstance(result, dict):
            chunk_findings = result.get("findings", [])
            if isinstance(chunk_findings, list):
                for finding in chunk_findings:
                    if isinstance(finding, dict) and finding.get("original_text"):
                        finding["chunk"] = chunk_idx + 1
                        findings.append(finding)
        return chunk_idx, findings

    all_findings = []
    with ThreadPoolExecutor(max_workers=min(max_workers, total_chunks)) as executor:
        futures = {
            executor.submit(process_chunk, idx, chunk_text): idx
            for idx, (char_offset, chunk_text) in enumerate(chunks)
        }
        for future in as_completed(futures):
            try:
                chunk_idx, findings = future.result()
                all_findings.extend(findings)
            except Exception as e:
                import logging
                logging.warning(f"Writing analysis chunk failed: {str(e)}")

    all_findings.sort(key=lambda f: f.get("chunk", 0))

    seen_texts = set()
    deduplicated = []
    for finding in all_findings:
        orig = finding.get("original_text", "").strip().lower()
        if orig and orig not in seen_texts:
            seen_texts.add(orig)
            deduplicated.append(finding)

    def _normalize_for_compare(s):
        s = s.replace(" ", "").replace("\u00a0", "").replace("\u200b", "")
        s = s.replace("\u2013", "-").replace("\u2014", "-").replace("\u2015", "-")
        s = s.replace("\u2018", "'").replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
        return s.lower()

    filtered = []
    for finding in deduplicated:
        orig = finding.get("original_text", "").strip()
        suggestion = finding.get("suggestion", "").strip()

        if _normalize_for_compare(orig) == _normalize_for_compare(suggestion) and orig != suggestion:
            continue

        orig_collapsed = re.sub(r'\s+', '', orig).lower()
        sugg_collapsed = re.sub(r'\s+', '', suggestion).lower()
        if orig_collapsed == sugg_collapsed and orig != suggestion:
            continue

        orig_no_hyphspace = orig.replace(" -", "-").replace("- ", "-").replace(" ", "")
        sugg_no_hyphspace = suggestion.replace(" -", "-").replace("- ", "-").replace(" ", "")
        if orig_no_hyphspace.lower() == sugg_no_hyphspace.lower() and orig != suggestion:
            continue

        if finding.get("type") == "Spelling":
            if orig_collapsed == sugg_collapsed:
                continue
            words_orig = orig.lower().split()
            words_sugg = suggestion.lower().split()
            if len(words_orig) > len(words_sugg) and len(words_sugg) >= 1:
                joined_check = "".join(words_orig)
                sugg_joined = "".join(words_sugg)
                if joined_check == sugg_joined:
                    continue

            if orig_collapsed.isalpha() and _is_valid_english_word(orig_collapsed):
                continue

            if orig.lower() in _get_common_words():
                continue

            if not _validate_spelling_finding(finding):
                continue

        filtered.append(finding)

    type_order = {"High": 0, "Medium": 1, "Low": 2}
    filtered.sort(key=lambda f: type_order.get(f.get("severity", "Low"), 2))

    return filtered


def create_gauge_chart(score, title, max_value=100):
    """
    Creates a gauge chart for a score
    
    Args:
        score (float): Score value
        title (str): Title for the gauge
        max_value (float): Maximum value for the gauge (default 100 for percentage)
        
    Returns:
        plotly figure object
    """
    # Define color thresholds based on percentage
    threshold_low = max_value * 0.4
    threshold_mid = max_value * 0.7
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        title={'text': title},
        domain={'x': [0, 1], 'y': [0, 1]},
        number={'suffix': '%' if max_value == 100 else ''},
        gauge={
            'axis': {'range': [0, max_value], 'tickwidth': 1, 'tickcolor': "darkblue"},
            'bar': {'color': "royalblue"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, threshold_low], 'color': 'lightcoral'},
                {'range': [threshold_low, threshold_mid], 'color': 'lightyellow'},
                {'range': [threshold_mid, max_value], 'color': 'lightgreen'}
            ]
        }
    ))
    
    fig.update_layout(
        height=300,
        width=400,
        margin=dict(l=30, r=30, b=20, t=40),
        paper_bgcolor="white",
        font={'color': "darkblue", 'family': "Arial"}
    )
    
    return fig
    
def sanitize_text_for_pdf(text):
    """
    Sanitize text to be compatible with FPDF's latin-1 encoding
    
    Args:
        text (str): Text to sanitize
        
    Returns:
        str: Sanitized text
    """
    if not text:
        return ""
    
    # Replace problematic characters
    replacements = {
        # Quotes and apostrophes
        '\u2018': "'",  # Left single quotation mark
        '\u2019': "'",  # Right single quotation mark
        '\u201c': '"',  # Left double quotation mark
        '\u201d': '"',  # Right double quotation mark
        
        # Dashes and hyphens
        '\u2013': '-',  # En dash
        '\u2014': '--', # Em dash
        '\u2015': '--', # Horizontal bar
        '\u2212': '-',  # Minus sign
        
        # Bullets and list markers
        '\u2022': '-',  # Bullet
        '\u2023': '-',  # Triangular bullet
        '\u2043': '-',  # Hyphen bullet
        '\u25e6': '-',  # White bullet
        '\u2219': '-',  # Bullet operator
        
        # Spaces and breaks
        '\u00a0': ' ',  # Non-breaking space
        '\u2000': ' ',  # En quad
        '\u2001': ' ',  # Em quad
        '\u2002': ' ',  # En space
        '\u2003': ' ',  # Em space
        '\u2004': ' ',  # Three-per-em space
        '\u2005': ' ',  # Four-per-em space
        '\u2006': ' ',  # Six-per-em space
        '\u2007': ' ',  # Figure space
        '\u2008': ' ',  # Punctuation space
        '\u2009': ' ',  # Thin space
        '\u200a': ' ',  # Hair space
        
        # Other punctuation
        '\u2026': '...', # Ellipsis
        '\u00b7': '.',   # Middle dot
        
        # Symbols
        '\u2713': 'v',   # Check mark
        '\u2192': '->',  # Rightward arrow
        '\u21e8': '->',  # Rightward white arrow
        '\u25ba': '>'    # Black right-pointing pointer
    }
    
    for original, replacement in replacements.items():
        text = text.replace(original, replacement)
    
    # Remove any other non-latin1 characters
    text = text.encode('latin-1', errors='replace').decode('latin-1')
    
    return text

def generate_report_pdf(filename, document_name, document_summary, assessment_results, assessment_areas, assessment_criteria, recommendations="", weighted_scores=None, competency_mapping=None, tags_data=None, writing_findings=None):
    """
    Generate a PDF report of the case study assessment
    
    Args:
        filename (str): Name to save the PDF as
        document_name (str): Name of the document
        document_summary (str): Summary of the document
        assessment_results (dict): Results of the assessment
        assessment_areas (dict): Assessment areas information
        assessment_criteria (dict): Assessment criteria information
        recommendations (str, optional): Recommendations for improvement
        weighted_scores (dict, optional): Weighted score calculations
        competency_mapping (dict, optional): KCM competency mapping with behavioral and functional competencies
        
    Returns:
        BytesIO: PDF file as BytesIO object
    """
    # Sanitize all text inputs for PDF compatibility
    document_name = sanitize_text_for_pdf(document_name)
    document_summary = sanitize_text_for_pdf(document_summary)
    if recommendations:
        recommendations = sanitize_text_for_pdf(recommendations)
    
    class PDF(FPDF):
        def header(self):
            # Set up the header
            self.set_font('Arial', 'B', 15)
            self.cell(0, 10, 'Case Study Assessment Report', 0, 1, 'C')
            self.ln(5)
        
        def footer(self):
            # Set up the footer
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            date = datetime.datetime.now().strftime("%Y-%m-%d")
            self.cell(0, 10, f'Generated on {date} - Page {self.page_no()}', 0, 0, 'C')
        
        def section_title(self, title):
            # Add a section title
            self.set_font('Arial', 'B', 14)
            self.set_fill_color(230, 230, 250)  # Light purple background
            self.cell(0, 10, title, 0, 1, 'L', True)
            self.ln(5)
        
        def sub_section_title(self, title):
            # Add a subsection title
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, title, 0, 1, 'L')
            self.ln(2)
        
        def add_score_box(self, label, score, max_score=3):
            # Create a visual score box
            self.set_fill_color(240, 240, 240)
            self.set_font('Arial', 'B', 10)
            self.cell(40, 10, label, 1, 0, 'L')
            
            # Choose color based on score percentage
            score_pct = score / max_score if max_score > 0 else 0
            if score_pct >= 0.75:
                self.set_fill_color(144, 238, 144)  # Light green
            elif score_pct >= 0.5:
                self.set_fill_color(255, 255, 153)  # Light yellow
            else:
                self.set_fill_color(255, 204, 203)  # Light red
                
            self.cell(20, 10, f"{score}/{max_score}", 1, 1, 'C', True)
            self.ln(1)
    
    # Create PDF object
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Cover page with title and document info
    pdf.set_font('Arial', 'B', 20)
    pdf.cell(0, 20, 'Case Study Assessment Report', 0, 1, 'C')
    pdf.ln(10)
    
    pdf.set_font('Arial', 'B', 14)
    pdf.cell(0, 10, f"Document: {document_name}", 0, 1, 'C')
    pdf.ln(10)
    
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 0, 1, 'C')
    
    # Add Final Composite Score prominently if available
    if weighted_scores:
        pdf.ln(15)
        final_score = weighted_scores.get("final_composite_score", 0)
        
        # Determine grade label
        if final_score >= 85:
            grade_label = "Excellent"
            grade_color = (40, 167, 69)  # Green
        elif final_score >= 70:
            grade_label = "Good"
            grade_color = (59, 130, 246)  # Blue
        elif final_score >= 55:
            grade_label = "Satisfactory"
            grade_color = (255, 193, 7)  # Yellow
        elif final_score >= 40:
            grade_label = "Needs Improvement"
            grade_color = (253, 126, 20)  # Orange
        else:
            grade_label = "Poor"
            grade_color = (220, 53, 69)  # Red
        
        pdf.set_fill_color(*grade_color)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('Arial', 'B', 18)
        pdf.cell(0, 15, f"Final Composite Score: {final_score:.1f}% - {grade_label}", 0, 1, 'C', True)
        pdf.set_text_color(0, 0, 0)  # Reset to black
    
    # Add logos side by side below the Final Composite Score
    logo_width = 35
    page_width = 210  # A4 width in mm
    total_logo_width = logo_width * 2 + 20  # Two logos with gap
    start_x = (page_width - total_logo_width) / 2
    
    cbc_logo_path = "attached_assets/cbc_logo_1770210514857.png"
    agk_logo_path = "attached_assets/agk_logo_1770210514857.png"
    
    if os.path.exists(cbc_logo_path) and os.path.exists(agk_logo_path):
        pdf.ln(15)
        current_y = pdf.get_y()
        pdf.image(cbc_logo_path, x=start_x, y=current_y, w=logo_width)
        pdf.image(agk_logo_path, x=start_x + logo_width + 20, y=current_y, w=logo_width)
        pdf.ln(40)
    
    # Executive summary page
    pdf.add_page()
    pdf.section_title("Executive Summary")
    pdf.set_font('Arial', '', 11)
    pdf.multi_cell(0, 6, document_summary)
    pdf.ln(5)
    
    if tags_data and isinstance(tags_data, dict):
        sectors = tags_data.get("sectors", [])
        subthemes = tags_data.get("subthemes", {})
        keywords = tags_data.get("keywords", [])
        
        if sectors or keywords:
            pdf.section_title("Sector Classification & Keywords")
            
            if sectors:
                pdf.sub_section_title("Applicable Sectors:")
                for sector in sectors:
                    sector_name = sanitize_text_for_pdf(str(sector))
                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_fill_color(30, 58, 138)
                    pdf.set_text_color(255, 255, 255)
                    pdf.cell(0, 8, f"  {sector_name}", 0, 1, 'L', True)
                    pdf.set_text_color(0, 0, 0)
                    
                    sector_subthemes = subthemes.get(sector, [])
                    if sector_subthemes:
                        pdf.set_font('Arial', '', 9)
                        for sub in sector_subthemes:
                            sub_name = sanitize_text_for_pdf(str(sub))
                            pdf.set_fill_color(240, 240, 250)
                            pdf.cell(10, 6, "", 0, 0)
                            pdf.cell(0, 6, f"- {sub_name}", 0, 1, 'L', True)
                    
                    pdf.ln(2)
                pdf.ln(3)
            
            if keywords:
                pdf.sub_section_title("Keywords:")
                pdf.set_font('Arial', '', 10)
                sanitized_keywords = [sanitize_text_for_pdf(str(k)) for k in keywords]
                keywords_text = ", ".join(sanitized_keywords)
                pdf.multi_cell(0, 6, keywords_text)
                pdf.ln(3)
            
            pdf.ln(5)
    
    # KCM Competency Mapping section
    if competency_mapping and isinstance(competency_mapping, dict):
        pdf.section_title("KCM Competency Mapping")
        pdf.set_font('Arial', '', 10)
        pdf.multi_cell(0, 5, "Based on the case study analysis, the following Karmayogi Competency Model (KCM) competencies have been identified:")
        pdf.ln(3)
        
        # Behavioral Competencies
        behavioral = competency_mapping.get("behavioral_competencies", [])
        if behavioral:
            pdf.set_font('Arial', 'B', 11)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 8, "Behavioral Sub-competencies", 0, 1, 'L', True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(2)
            
            for i, comp in enumerate(behavioral, 1):
                if isinstance(comp, dict):
                    comp_name = sanitize_text_for_pdf(comp.get("name", "Unknown"))
                    parent = sanitize_text_for_pdf(comp.get("parent_competency", ""))
                    display_name = f"{comp_name} ({parent})" if parent else comp_name
                    justification = sanitize_text_for_pdf(comp.get("justification", "No justification provided"))
                    
                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_fill_color(240, 240, 240)
                    pdf.cell(0, 7, f"{i}. {display_name}", 0, 1, 'L', True)
                    
                    pdf.set_font('Arial', '', 9)
                    pdf.multi_cell(0, 5, f"   {justification}")
                    pdf.ln(2)
            pdf.ln(3)
        
        # Functional Competencies
        functional = competency_mapping.get("functional_competencies", [])
        if functional:
            pdf.set_font('Arial', 'B', 11)
            pdf.set_fill_color(30, 58, 138)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(0, 8, "Functional Sub-competencies", 0, 1, 'L', True)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(2)
            
            for i, comp in enumerate(functional, 1):
                if isinstance(comp, dict):
                    comp_name = sanitize_text_for_pdf(comp.get("name", "Unknown"))
                    parent = sanitize_text_for_pdf(comp.get("parent_competency", ""))
                    display_name = f"{comp_name} ({parent})" if parent else comp_name
                    justification = sanitize_text_for_pdf(comp.get("justification", "No justification provided"))
                    
                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_fill_color(240, 240, 240)
                    pdf.cell(0, 7, f"{i}. {display_name}", 0, 1, 'L', True)
                    
                    pdf.set_font('Arial', '', 9)
                    pdf.multi_cell(0, 5, f"   {justification}")
                    pdf.ln(2)
            pdf.ln(3)
        
        pdf.ln(5)
    
    # Weighted scoring summary
    if weighted_scores:
        pdf.section_title("Weighted Scoring Summary")
        
        # Create a table for weighted scores
        pdf.set_fill_color(240, 240, 240)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(70, 8, "Assessment Area", 1, 0, 'L', True)
        pdf.cell(25, 8, "Score", 1, 0, 'C', True)
        pdf.cell(25, 8, "Max", 1, 0, 'C', True)
        pdf.cell(25, 8, "%", 1, 0, 'C', True)
        pdf.cell(25, 8, "Weight", 1, 0, 'C', True)
        pdf.cell(25, 8, "Contrib.", 1, 1, 'C', True)
        
        pdf.set_font('Arial', '', 10)
        for area_id, area_info in assessment_areas.items():
            area_score_data = weighted_scores["area_scores"].get(area_id, {})
            weighted_data = weighted_scores["weighted_scores"].get(area_id, {})
            
            score = area_score_data.get("score", 0)
            max_score = area_score_data.get("max_score", 1)
            percentage = area_score_data.get("percentage", 0)
            weight = weighted_data.get("weight", 0) * 100
            contribution = weighted_data.get("weighted_contribution", 0) * 100
            
            # Shorten long area names
            area_name = area_info["name"]
            if len(area_name) > 35:
                area_name = area_name[:32] + "..."
            
            # Color based on percentage
            if percentage >= 75:
                pdf.set_fill_color(220, 255, 220)
            elif percentage >= 50:
                pdf.set_fill_color(255, 255, 220)
            else:
                pdf.set_fill_color(255, 220, 220)
            
            pdf.cell(70, 8, area_name, 1, 0, 'L')
            pdf.cell(25, 8, f"{score}", 1, 0, 'C')
            pdf.cell(25, 8, f"{max_score}", 1, 0, 'C')
            pdf.cell(25, 8, f"{percentage:.1f}%", 1, 0, 'C', True)
            pdf.cell(25, 8, f"{weight:.0f}%", 1, 0, 'C')
            pdf.cell(25, 8, f"{contribution:.1f}%", 1, 1, 'C')
        
        pdf.ln(5)
        
        # Final score row
        final_score = weighted_scores.get("final_composite_score", 0)
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(145, 10, "Final Composite Score:", 1, 0, 'R')
        
        if final_score >= 75:
            pdf.set_fill_color(144, 238, 144)
        elif final_score >= 50:
            pdf.set_fill_color(255, 255, 153)
        else:
            pdf.set_fill_color(255, 204, 203)
        
        pdf.cell(50, 10, f"{final_score:.1f}%", 1, 1, 'C', True)
        pdf.ln(10)
    
    # Detailed assessment by area
    for area_id, area_info in assessment_areas.items():
        pdf.add_page()
        pdf.section_title(f"Area: {area_info['name']}")
        
        pdf.set_font('Arial', '', 10)
        pdf.multi_cell(0, 5, area_info["description"])
        pdf.ln(3)
        
        # Display area weight and total points
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, f"Weight: {area_info.get('weight', 0) * 100:.0f}% | Total Points: {area_info.get('total_points', 0)}", 0, 1)
        pdf.ln(5)
        
        if area_id in assessment_results:
            area_results = assessment_results[area_id]
            
            # Calculate area scores (skip informational criteria)
            area_criteria_defs = assessment_criteria.get(area_id, {})
            total_score = sum(
                result.get("score", 0) for crit_id, result in area_results.items()
                if not area_criteria_defs.get(crit_id, {}).get("informational", False)
            )
            max_possible = area_info.get("total_points", 1)
            percentage = (total_score / max_possible * 100) if max_possible > 0 else 0
            
            # Display area scores in a nice box with a colored background
            if percentage >= 75:
                pdf.set_fill_color(200, 255, 200)  # Light green
            elif percentage >= 50:
                pdf.set_fill_color(255, 255, 200)  # Light yellow
            else:
                pdf.set_fill_color(255, 220, 220)  # Light red
                
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(190, 12, f"Area Score: {total_score}/{max_possible} ({percentage:.1f}%)", 1, 1, 'C', True)
            pdf.ln(5)
            
            # Create a bar chart visualization for the PDF
            pdf.sub_section_title("Criteria Score Visualization")
            
            try:
                # Create a simple visualization using the FPDF directly for reliability
                area_criteria_data = assessment_criteria[area_id]
                
                # Get criteria names and scores
                criteria_list = list(area_results.keys())
                
                # Set up the chart dimensions
                chart_x = 55  # Increased to give more room for labels
                max_bar_width = 100  # Maximum width of bars
                bar_height = 8
                space_between = 5
                header_height = 20
                
                # Calculate total chart height
                chart_height = header_height + (len(criteria_list) * (bar_height + space_between))
                
                # Check if we need a new page
                if pdf.get_y() + chart_height > pdf.page_break_trigger:
                    pdf.add_page()
                
                # Draw chart header
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 10, f"Criteria Scores", 0, 1, 'C') # Centered header
                
                pdf.ln(2)
                
                # Draw each criterion score as a bar (skip informational criteria)
                for crit_id in criteria_list:
                    crit_info = area_criteria_data[crit_id]
                    if crit_info.get("informational", False):
                        continue
                    criterion_name = crit_info["name"]
                    max_score = crit_info.get("max_score", 3)
                    score = area_results[crit_id].get("score", 0)
                    
                    # Draw criterion name
                    bar_y = pdf.get_y()
                    
                    # Position the text before the chart
                    pdf.set_font('Arial', '', 8) # Smaller font for more room
                    
                    # No longer truncating as aggressively, using 45 chars limit
                    if len(criterion_name) > 45:
                        criterion_name = criterion_name[:42] + "..."
                    
                    pdf.set_xy(10, bar_y) # Start from left margin
                    pdf.cell(chart_x - 12, bar_height, criterion_name, 0, 0, 'R')
                    
                    # Draw bar background (gray)
                    pdf.set_fill_color(240, 240, 240)
                    pdf.rect(chart_x, bar_y, max_bar_width, bar_height, 'F')
                    
                    # Choose bar color based on score percentage
                    score_pct = score / max_score if max_score > 0 else 0
                    if score_pct >= 0.75:
                        pdf.set_fill_color(144, 238, 144)  # Light green
                    elif score_pct >= 0.5:
                        pdf.set_fill_color(255, 255, 153)  # Light yellow
                    else:
                        pdf.set_fill_color(255, 204, 203)  # Light red
                    
                    # Draw the score bar
                    bar_width = score_pct * max_bar_width
                    if bar_width > 0:
                        pdf.rect(chart_x, bar_y, bar_width, bar_height, 'F')
                    
                    # Add a black border around the bar
                    pdf.set_draw_color(0, 0, 0)
                    pdf.rect(chart_x, bar_y, max_bar_width, bar_height, 'D')
                    
                    # Add score text
                    pdf.set_font('Arial', 'B', 8)
                    pdf.set_xy(chart_x + max_bar_width + 2, bar_y)
                    pdf.cell(15, bar_height, f"{score}/{max_score}", 0, 1, 'L')
                    
                    # Space for next bar
                    pdf.ln(space_between)
                
                pdf.ln(5)
                
            except Exception as e:
                # If visualization fails, just continue without the chart
                pdf.set_font('Arial', 'I', 10)
                pdf.cell(0, 10, f"Chart visualization not available: {str(e)}", 0, 1)
            
            # Process each criterion with detailed info
            area_criteria = assessment_criteria[area_id]
            
            pdf.sub_section_title("Detailed Criteria Assessment")
            
            for criterion_id, result in area_results.items():
                crit_info = area_criteria[criterion_id]
                criterion_name = crit_info["name"]
                is_informational = crit_info.get("informational", False)

                if is_informational:
                    narrative = result.get("narrative", result.get("reasoning", "No analysis provided"))
                    narrative = sanitize_text_for_pdf(narrative)

                    pdf.set_draw_color(59, 130, 246)
                    pdf.set_fill_color(219, 234, 254)
                    pdf.set_font('Arial', 'B', 11)
                    pdf.cell(0, 8, f"{sanitize_text_for_pdf(criterion_name)} (Informational)", 1, 1, 'L', True)

                    pdf.set_font('Arial', 'I', 9)
                    pdf.set_text_color(59, 130, 246)
                    pdf.cell(0, 6, "This criterion provides informational analysis and is not scored.", 0, 1)
                    pdf.set_text_color(0, 0, 0)

                    pdf.set_font('Arial', 'B', 10)
                    pdf.cell(0, 8, "Analysis:", 0, 1)
                    pdf.set_font('Arial', '', 10)
                    pdf.multi_cell(0, 6, narrative)

                    doc_ref = result.get("document_reference", "")
                    if doc_ref:
                        doc_ref = sanitize_text_for_pdf(doc_ref)
                        pdf.set_font('Arial', 'B', 10)
                        pdf.cell(0, 8, "Document References:", 0, 1)
                        pdf.set_font('Arial', 'I', 9)
                        pdf.multi_cell(0, 6, doc_ref)

                    pdf.set_draw_color(0, 0, 0)
                    pdf.ln(5)
                else:
                    max_score = crit_info.get("max_score", 3)
                    scoring_logic = crit_info.get("scoring_logic", "")
                    score = result.get("score", 0)
                    reasoning = result.get("reasoning", "No reasoning provided")
                    doc_ref = result.get("document_reference", "No references provided")

                    reasoning = sanitize_text_for_pdf(reasoning)
                    doc_ref = sanitize_text_for_pdf(doc_ref)
                    scoring_logic = sanitize_text_for_pdf(scoring_logic)

                    pdf.set_draw_color(100, 100, 100)
                    pdf.set_fill_color(240, 248, 255)
                    pdf.set_font('Arial', 'B', 11)
                    pdf.cell(0, 8, f"{sanitize_text_for_pdf(criterion_name)}", 1, 1, 'L', True)

                    pdf.set_font('Arial', 'I', 9)
                    pdf.cell(0, 6, f"Scoring: {scoring_logic}", 0, 1)

                    score_pct = score / max_score if max_score > 0 else 0
                    if score_pct >= 0.75:
                        pdf.set_fill_color(144, 238, 144)
                    elif score_pct >= 0.5:
                        pdf.set_fill_color(255, 255, 153)
                    else:
                        pdf.set_fill_color(255, 204, 203)

                    pdf.set_font('Arial', 'B', 10)
                    pdf.cell(30, 8, "Score:", 0, 0)
                    pdf.cell(20, 8, f"{score}/{max_score}", 1, 1, 'C', True)

                    pdf.set_font('Arial', 'B', 10)
                    pdf.cell(0, 8, "Reasoning & Evidence:", 0, 1)
                    pdf.set_font('Arial', '', 10)
                    pdf.multi_cell(0, 6, reasoning)

                    pdf.set_font('Arial', 'B', 10)
                    pdf.cell(0, 8, "Document References:", 0, 1)
                    pdf.set_font('Arial', 'I', 9)
                    pdf.multi_cell(0, 6, doc_ref)

                    pdf.ln(5)

            if area_id == "area2" and writing_findings and isinstance(writing_findings, list) and len(writing_findings) > 0:
                pdf.add_page()
                pdf.section_title("Writing Assistant Findings")
                pdf.set_font('Arial', '', 10)
                pdf.multi_cell(0, 5, "Detailed analysis of grammar, spelling, tense consistency, redundancy, and sentence structure across the complete case study.")
                pdf.ln(3)

                type_counts = {}
                severity_counts = {"High": 0, "Medium": 0, "Low": 0}
                for finding in writing_findings:
                    t = finding.get("type", "Other")
                    s = finding.get("severity", "Low")
                    type_counts[t] = type_counts.get(t, 0) + 1
                    if s in severity_counts:
                        severity_counts[s] += 1

                pdf.set_font('Arial', 'B', 10)
                pdf.cell(0, 8, f"Summary: {len(writing_findings)} issues found", 0, 1)
                pdf.set_font('Arial', '', 9)
                summary_parts = []
                for issue_type, count in type_counts.items():
                    summary_parts.append(f"{issue_type}: {count}")
                for sev, count in severity_counts.items():
                    if count > 0:
                        summary_parts.append(f"{sev} severity: {count}")
                pdf.cell(0, 6, " | ".join(summary_parts), 0, 1)
                pdf.ln(5)

                col_widths = [12, 20, 48, 48, 16, 46]
                headers = ["Issue #", "Type", "Original Text", "Suggested Fix", "Severity", "Explanation"]
                line_height = 4

                def draw_table_header():
                    pdf.set_fill_color(30, 58, 138)
                    pdf.set_text_color(255, 255, 255)
                    pdf.set_font('Arial', 'B', 7)
                    for j, header in enumerate(headers):
                        pdf.cell(col_widths[j], 7, header, 1, 0, 'C', True)
                    pdf.ln()
                    pdf.set_text_color(0, 0, 0)

                def calc_lines(text, col_width):
                    if not text:
                        return 1
                    pdf.set_font('Arial', '', 7)
                    char_width = pdf.get_string_width('A')
                    chars_per_line = max(1, int(col_width / char_width))
                    words = text.split()
                    lines = 1
                    current_line_len = 0
                    for word in words:
                        word_len = len(word)
                        if current_line_len == 0:
                            current_line_len = word_len
                        elif current_line_len + 1 + word_len <= chars_per_line:
                            current_line_len += 1 + word_len
                        else:
                            lines += 1
                            current_line_len = word_len
                    return lines

                def draw_wrapped_cell(x, y, w, h, text, align='L', fill=True):
                    pdf.set_xy(x, y)
                    pdf.rect(x, y, w, h, 'DF')
                    pdf.set_xy(x + 1, y + 1)
                    pdf.multi_cell(w - 2, line_height, text, 0, align)

                draw_table_header()

                pdf.set_font('Arial', '', 7)
                for idx, finding in enumerate(writing_findings, 1):
                    sev = finding.get("severity", "Low")
                    orig = sanitize_text_for_pdf(str(finding.get("original_text", "")))
                    sugg = sanitize_text_for_pdf(str(finding.get("suggestion", "")))
                    expl = sanitize_text_for_pdf(str(finding.get("context", "")))
                    f_type = sanitize_text_for_pdf(str(finding.get("type", "Other")))

                    max_lines = max(
                        calc_lines(orig, col_widths[2]),
                        calc_lines(sugg, col_widths[3]),
                        calc_lines(expl, col_widths[5]),
                        1
                    )
                    row_height = max(7, max_lines * line_height + 2)

                    if pdf.get_y() + row_height > pdf.page_break_trigger:
                        pdf.add_page()
                        draw_table_header()

                    if sev == "High":
                        pdf.set_fill_color(248, 215, 218)
                    elif sev == "Medium":
                        pdf.set_fill_color(255, 243, 205)
                    else:
                        pdf.set_fill_color(245, 245, 245)

                    row_y = pdf.get_y()
                    row_x = pdf.get_x()

                    pdf.set_font('Arial', '', 7)
                    x_pos = row_x
                    pdf.rect(x_pos, row_y, col_widths[0], row_height, 'DF')
                    pdf.set_xy(x_pos, row_y + (row_height - line_height) / 2)
                    pdf.cell(col_widths[0], line_height, str(idx), 0, 0, 'C')
                    x_pos += col_widths[0]

                    pdf.rect(x_pos, row_y, col_widths[1], row_height, 'DF')
                    pdf.set_xy(x_pos, row_y + (row_height - line_height) / 2)
                    pdf.cell(col_widths[1], line_height, f_type, 0, 0, 'C')
                    x_pos += col_widths[1]

                    draw_wrapped_cell(x_pos, row_y, col_widths[2], row_height, orig)
                    x_pos += col_widths[2]

                    if sev == "High":
                        pdf.set_fill_color(248, 215, 218)
                    elif sev == "Medium":
                        pdf.set_fill_color(255, 243, 205)
                    else:
                        pdf.set_fill_color(245, 245, 245)

                    draw_wrapped_cell(x_pos, row_y, col_widths[3], row_height, sugg)
                    x_pos += col_widths[3]

                    if sev == "High":
                        pdf.set_fill_color(248, 215, 218)
                    elif sev == "Medium":
                        pdf.set_fill_color(255, 243, 205)
                    else:
                        pdf.set_fill_color(245, 245, 245)

                    pdf.rect(x_pos, row_y, col_widths[4], row_height, 'DF')
                    pdf.set_xy(x_pos, row_y + (row_height - line_height) / 2)
                    pdf.cell(col_widths[4], line_height, sev, 0, 0, 'C')
                    x_pos += col_widths[4]

                    draw_wrapped_cell(x_pos, row_y, col_widths[5], row_height, expl)

                    pdf.set_y(row_y + row_height)

                pdf.ln(5)

        else:
            pdf.set_font('Arial', 'I', 10)
            pdf.cell(0, 10, "No assessment data available for this area", 0, 1)
    
    # Add recommendations section
    if recommendations:
        pdf.add_page()
        pdf.section_title("Recommendations for Improvement")
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 6, "Based on the assessment results, the following recommendations are provided to improve the case study:")
        pdf.ln(5)
        
        # Format the recommendations text - improved handling
        pdf.set_font('Arial', '', 11)
        
        # Process recommendations with Recommendation X: format
        if "Recommendation" in recommendations:
            # Split by "Recommendation" and process each item
            parts = recommendations.split("Recommendation")
            for part in parts:
                if part.strip():
                    # Clean up the format, preserving sentence integrity
                    cleaned_part = part.strip().replace('\n', ' ')
                    # Remove any stray bullets or formatting that might be causing breaks
                    cleaned_part = cleaned_part.replace('•', '').replace('- ', '')
                    # Sanitize for PDF compatibility
                    cleaned_part = sanitize_text_for_pdf(cleaned_part)
                    
                    # Handle recommendation number formatting
                    if ':' in cleaned_part and cleaned_part.find(':') < 10:
                        num, content = cleaned_part.split(':', 1)
                        
                        # Create a full recommendation text with proper formatting
                        pdf.set_font('Arial', 'B', 11)
                        pdf.write(6, f"Recommendation {num.strip()}: ")
                        
                        # Continue the same line with regular font for content
                        pdf.set_font('Arial', '', 11)
                        pdf.multi_cell(0, 6, content.strip())
                    else:
                        # If no colon format, just print the whole thing
                        pdf.set_font('Arial', 'B', 11)
                        pdf.write(6, "Recommendation: ")
                        pdf.set_font('Arial', '', 11)
                        pdf.multi_cell(0, 6, cleaned_part)
                    
                    pdf.ln(4)
                    
        # Split recommendations by dash list items as a fallback
        elif "-" in recommendations:
            rec_points = recommendations.split("-")
            for point in rec_points:
                if point.strip():
                    # Clean up and replace newlines with spaces for continuous text
                    clean_point = point.strip().replace('\n', ' ')
                    # Sanitize for PDF compatibility
                    clean_point = sanitize_text_for_pdf(clean_point)
                    pdf.set_font('Arial', '', 11)
                    # Use bullet point symbol with proper indentation
                    pdf.cell(5, 6, "-", 0, 0)
                    pdf.multi_cell(0, 6, clean_point)
                    pdf.ln(2)
        # Fallback in case bullet points are used
        elif "•" in recommendations:
            rec_points = recommendations.split("•")
            for point in rec_points:
                if point.strip():
                    # Clean up and replace newlines with spaces for continuous text
                    clean_point = point.strip().replace('\n', ' ')
                    # Sanitize for PDF compatibility
                    clean_point = sanitize_text_for_pdf(clean_point)
                    pdf.set_font('Arial', '', 11)
                    # Use bullet point symbol with proper indentation
                    pdf.cell(5, 6, "-", 0, 0)
                    pdf.multi_cell(0, 6, clean_point)
                    pdf.ln(2)
        else:
            # Just show as regular text if no structured format detected
            # Sanitize for PDF compatibility
            sanitized_recommendations = sanitize_text_for_pdf(recommendations.replace('\n', ' '))
            pdf.multi_cell(0, 6, sanitized_recommendations)
    
    # Save to BytesIO
    pdf_bytes = pdf.output(dest='S')
    if isinstance(pdf_bytes, str):
        pdf_bytes = pdf_bytes.encode('latin-1')
    
    pdf_output = BytesIO(pdf_bytes)
    pdf_output.seek(0)
    
    return pdf_output

def get_download_link(pdf_bytes, filename, text):
    """
    Generate a download link for a PDF file
    
    Args:
        pdf_bytes (BytesIO): PDF file as BytesIO object
        filename (str): Filename for download
        text (str): Link text
        
    Returns:
        str: HTML for download link
    """
    b64 = base64.b64encode(pdf_bytes.read()).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="{filename}">{text}</a>'
    return href


def _extract_case_summary(case_text, max_chars=600):
    if not case_text or len(case_text) < 50:
        return ""
    paragraphs = [p.strip() for p in case_text.split('\n') if len(p.strip()) > 40]
    key_sentences = []
    for p in paragraphs[:20]:
        sentences = re.split(r'(?<=[.!?])\s+', p)
        for s in sentences:
            s = s.strip()
            if len(s) > 30 and any(kw in s.lower() for kw in [
                'challenge', 'initiative', 'implement', 'transform', 'innovat',
                'leader', 'district', 'mission', 'project', 'programme', 'policy',
                'reform', 'solution', 'strategy', 'impact', 'outcome', 'result',
                'governance', 'citizen', 'community', 'development', 'digital',
                'health', 'education', 'infrastructure', 'environment', 'technology'
            ]):
                key_sentences.append(s)
            if len(' '.join(key_sentences)) > max_chars:
                break
        if len(' '.join(key_sentences)) > max_chars:
            break
    summary = ' '.join(key_sentences)[:max_chars]
    if not summary and paragraphs:
        summary = ' '.join(paragraphs[:3])[:max_chars]
    return summary


def load_case_database(file_path):
    import pandas as pd

    try:
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            df = pd.read_excel(file_path)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            return [], "", 0
    except Exception as e:
        return [], "", 0

    col_names = list(df.columns)
    has_case_text = len(col_names) >= 3
    has_igot_link = len(col_names) >= 4

    cases = []
    for _, row in df.iterrows():
        title = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        desc = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        case_text = str(row.iloc[2]).strip() if has_case_text and pd.notna(row.iloc[2]) else ""
        igot_link = str(row.iloc[3]).strip() if has_igot_link and pd.notna(row.iloc[3]) else ""
        title = re.sub(r'\s+', ' ', title).strip()
        desc = re.sub(r'\s+', ' ', desc).strip()
        if igot_link and not igot_link.startswith('http'):
            igot_link = ""
        if title and desc:
            cases.append({
                "title": title,
                "description": desc,
                "case_text": case_text,
                "igot_link": igot_link
            })

    case_count = len(cases)
    structured_text = ""
    for idx, case in enumerate(cases, 1):
        entry = f"Case {idx}: {case['title']}\nDescription: {case['description']}\n"
        if case['case_text']:
            summary = _extract_case_summary(case['case_text'])
            if summary:
                entry += f"Key Details: {summary}\n"
        structured_text += entry + "\n"

    return cases, structured_text, case_count


def run_caseconnect_analysis(questionnaire, curriculum_text, case_database_text, case_count):
    learners = questionnaire.get("learners", "")
    objective = questionnaire.get("objective", "")
    competencies = questionnaire.get("competencies", [])
    duration = questionnaire.get("duration", "")
    sector = questionnaire.get("sector", "")

    competencies_str = ", ".join(competencies) if competencies else "Not specified"

    curriculum_section = ""
    if curriculum_text:
        curriculum_section = f"""
=== UPLOADED COURSE OUTLINE / CURRICULUM ===
{curriculum_text[:8000]}
===

The faculty has uploaded their course outline above. Analyse it to identify modules, themes, and learning objectives.
Use this to provide module-level case suggestions in addition to overall recommendations.
"""

    prompt = f"""You are CaseConnect, an expert AI case discovery tool for the Amrit Gyaan Kosh (AGK) repository of Indian governance case studies hosted on iGOT Karmayogi platform.

A faculty member is looking for case studies to integrate into their teaching programme. You have deep knowledge of each case's content, characters, challenges, and outcomes from the database below. Use this knowledge to make highly specific, grounded recommendations.

=== FACULTY INPUTS ===
1. Who are the learners? {learners}
2. Primary learning objective: {objective}
3. Competencies being focused on (Karmayogi Competency Model): {competencies_str}
4. Duration of the session: {duration}
5. Sector/Theme: {sector}

{curriculum_section}

=== AGK CASE STUDY DATABASE ({case_count} cases) ===
{case_database_text[:100000]}
===

Based on the above, provide your recommendations as a JSON object with this structure:
{{
    "case_recommendations": [
        {{
            "case_title": "Exact title from the database — must match precisely",
            "why_it_fits": "3-4 sentence explanation that references SPECIFIC elements from the case: name the protagonist/decision-maker, describe the actual challenge they faced, explain what governance approach or intervention was used, and connect this directly to the faculty's stated learning objective and learner profile. Do NOT write generic statements like 'this case aligns with the competency'. Instead explain the concrete teaching value.",
            "relevant_competencies": ["list of specific KCM competencies this case develops — explain the link"],
            "key_themes": ["2-4 governance themes this case covers, e.g. 'Digital Governance', 'Public Health', 'Citizen Engagement', 'Data-Driven Decision Making'"],
            "discussion_points": ["2-3 specific discussion questions for the classroom that draw from the actual case content, e.g. 'How did Sampath Kumar balance centralized monitoring with local autonomy in Meghalaya?'"],
            "suggested_duration": "recommended session duration for this case"
        }}
    ],
    "module_suggestions": [
        {{
            "module_name": "Module name from the curriculum (if provided)",
            "recommended_cases": [
                {{
                    "case_title": "Case title",
                    "relevance": "Specific reason why this case fits this module, referencing actual case content"
                }}
            ]
        }}
    ],
    "teaching_strategy": "A detailed paragraph suggesting how to best use these cases in the teaching programme. Include specific session format (e.g. role-play, group discussion, case analysis), discussion approach, sequencing advice for using multiple cases, and how to connect the cases to the competency development goals.",
    "additional_notes": "Any additional observations or suggestions for the faculty."
}}

Rules:
- Recommend 5-8 cases that best match the faculty's needs
- ONLY recommend cases that exist in the database above — do not invent or modify case titles
- The case_title MUST be an exact match to a title in the database
- In "why_it_fits", always reference the actual protagonist, situation, or outcome from the case — never write generic justifications
- Discussion points must be grounded in the specific case content, not generic governance questions
- Key themes should reflect the actual governance domains covered in the case
- If a curriculum was uploaded, provide module_suggestions mapping cases to specific modules
- If no curriculum was uploaded, return an empty list for module_suggestions
- Teaching strategy should be practical, actionable, and specific to the recommended cases
- Consider the session duration when making recommendations
- Prioritise cases that enable active learning (debate, role-play, problem-solving) over passive reading"""

    result = call_openai_api(prompt, response_format="json_object", temperature=0.2)

    if isinstance(result, str):
        return {
            "case_recommendations": [],
            "module_suggestions": [],
            "teaching_strategy": "Unable to generate recommendations. Please try again.",
            "additional_notes": str(result)
        }

    if not isinstance(result, dict):
        return {
            "case_recommendations": [],
            "module_suggestions": [],
            "teaching_strategy": "Unable to generate recommendations. Please try again.",
            "additional_notes": ""
        }

    if "case_recommendations" not in result:
        result["case_recommendations"] = []
    if "module_suggestions" not in result:
        result["module_suggestions"] = []
    if "teaching_strategy" not in result:
        result["teaching_strategy"] = ""
    if "additional_notes" not in result:
        result["additional_notes"] = ""

    return result
