"""Google Gemini provider for the CBC-India AGK Case Study Suite.

This module is an *optional* drop-in alternative to the OpenAI backend. It is
only used when the environment variable ``USE_GEMINI`` is truthy
(``1``/``true``/``yes``/``on``); otherwise the application keeps calling OpenAI
exactly as before and nothing in this file is imported.

Design notes
------------
* All AI traffic in the app funnels through ``utils.call_openai_api``. That
  wrapper decides the provider and, for Gemini, delegates to
  :func:`generate_text` here. The JSON post-processing, defaults and error
  fallbacks stay in ``utils`` so both providers behave identically to callers.
* :func:`generate_text` returns **plain text**, mirroring
  ``response.choices[0].message.content`` from the OpenAI SDK, and raises on
  failure (``utils`` converts that into the usual error string).
* Both Vertex AI (service-account) and Gemini Developer API (API key) auth are
  supported; Vertex AI is used when ``GOOGLE_PROJECT_ID`` is set.

Environment variables
---------------------
``USE_GEMINI``                    enable this provider (default: off -> OpenAI)
``GOOGLE_PROJECT_ID``             Vertex AI project (alias: GOOGLE_CLOUD_PROJECT)
``GOOGLE_LOCATION``               Vertex AI location (alias: GOOGLE_CLOUD_LOCATION, default: global)
``GOOGLE_APPLICATION_CREDENTIALS`` service-account JSON path (read by google-auth)
``GEMINI_API_KEY``/``GOOGLE_API_KEY``  used instead of Vertex AI when no project is set
``GENAI_MODEL_NAME``              model id (default: gemini-3.1-flash-lite)
``GENAI_MAX_OUTPUT_TOKENS``       output-token cap (default: 8192)
``GENAI_THINKING_LEVEL``          low | high | default (default: low)
"""

import json
import os
import re
import threading

# OpenAI caps responses at 2000 tokens. Gemini bills "thinking" tokens against
# max_output_tokens, so a like-for-like cap would truncate answers before any
# visible text is produced — hence the higher default.
DEFAULT_MODEL_NAME = "gemini-3.1-flash-lite"
DEFAULT_MAX_OUTPUT_TOKENS = 8192
DEFAULT_THINKING_LEVEL = "low"

_TRUTHY = {"1", "true", "yes", "on", "y"}

# Model ids handled by the OpenAI backend. Callers pass these as defaults
# (e.g. "gpt-4o"), so in Gemini mode they are replaced by GENAI_MODEL_NAME.
_OPENAI_MODEL_PREFIXES = ("gpt-", "gpt4", "o1", "o3", "o4", "chatgpt", "text-", "davinci")

_client = None
_client_lock = threading.Lock()


def use_gemini():
    """True when the app should route AI calls to Gemini instead of OpenAI."""
    return os.getenv("USE_GEMINI", "").strip().lower() in _TRUTHY


def get_model_name(requested_model=None):
    """Resolve the Gemini model id to use.

    An OpenAI-style model id (or nothing) resolves to ``GENAI_MODEL_NAME``; an
    explicit Gemini model id is passed through untouched.
    """
    configured = (os.getenv("GENAI_MODEL_NAME") or "").strip() or DEFAULT_MODEL_NAME
    if not requested_model:
        return configured
    requested = str(requested_model).strip()
    if requested.lower().startswith(_OPENAI_MODEL_PREFIXES):
        return configured
    return requested or configured


def _max_output_tokens():
    raw = (os.getenv("GENAI_MAX_OUTPUT_TOKENS") or "").strip()
    if raw.isdigit() and int(raw) > 0:
        return int(raw)
    return DEFAULT_MAX_OUTPUT_TOKENS


def get_client():
    """Return a lazily created, process-wide ``genai.Client``.

    The client is thread-safe and reused across the app's ThreadPoolExecutor
    workers (see ``utils.analyze_writing_quality_chunked``).
    """
    global _client
    if _client is not None:
        return _client
    with _client_lock:
        if _client is not None:
            return _client
        try:
            from google import genai
        except ImportError as exc:  # pragma: no cover - depends on install
            raise RuntimeError(
                "USE_GEMINI is enabled but the google-genai package is not "
                "installed. Install it with: pip install google-genai"
            ) from exc

        project = (os.getenv("GOOGLE_PROJECT_ID") or os.getenv("GOOGLE_CLOUD_PROJECT") or "").strip()
        location = (
            os.getenv("GOOGLE_LOCATION") or os.getenv("GOOGLE_CLOUD_LOCATION") or "global"
        ).strip()
        api_key = (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()

        if project:
            _client = genai.Client(vertexai=True, project=project, location=location)
        elif api_key:
            _client = genai.Client(api_key=api_key)
        else:
            raise RuntimeError(
                "USE_GEMINI is enabled but no Google credentials were found. "
                "Set GOOGLE_PROJECT_ID (+ GOOGLE_APPLICATION_CREDENTIALS) for "
                "Vertex AI, or GEMINI_API_KEY for the Gemini Developer API."
            )
        return _client


def describe_provider():
    """One-line description of the active Gemini configuration, for startup logs.

    Never includes the API key or any credential contents — only whether a
    credential is present and which auth mode was selected.
    """
    project = (os.getenv("GOOGLE_PROJECT_ID") or os.getenv("GOOGLE_CLOUD_PROJECT") or "").strip()
    location = (
        os.getenv("GOOGLE_LOCATION") or os.getenv("GOOGLE_CLOUD_LOCATION") or "global"
    ).strip()
    if project:
        credentials = (os.getenv("GOOGLE_APPLICATION_CREDENTIALS") or "").strip()
        auth = "Vertex AI (project=%s, location=%s, credentials=%s)" % (
            project, location, "set" if credentials else "ADC/default",
        )
    elif (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip():
        auth = "Gemini Developer API (api key set)"
    else:
        auth = "NO CREDENTIALS FOUND"
    level = (os.getenv("GENAI_THINKING_LEVEL") or DEFAULT_THINKING_LEVEL).strip().lower()
    return "Gemini | model=%s | auth=%s | max_output_tokens=%d | thinking_level=%s" % (
        get_model_name(), auth, _max_output_tokens(), level,
    )


def _thinking_config(types):
    """Build a ThinkingConfig from GENAI_THINKING_LEVEL, or None for defaults.

    ``thinking_level`` is the Gemini 3 control; older families use
    ``thinking_budget`` and reject it, so callers retry without any thinking
    config when the API complains (see ``_is_thinking_unsupported``).
    """
    level = (os.getenv("GENAI_THINKING_LEVEL") or DEFAULT_THINKING_LEVEL).strip().lower()
    if level in ("low", "high"):
        try:
            return types.ThinkingConfig(thinking_level=level)
        except Exception:
            return None
    return None


def _build_config(types, response_format, temperature, seed, max_output_tokens, thinking=None):
    kwargs = {
        # Gemini accepts 0.0-2.0, same as OpenAI's usable range.
        "temperature": max(0.0, min(2.0, float(temperature))),
        "max_output_tokens": max_output_tokens,
    }
    if response_format == "json_object":
        # Gemini's equivalent of OpenAI's response_format={"type": "json_object"}.
        kwargs["response_mime_type"] = "application/json"
    if seed is not None:
        kwargs["seed"] = int(seed)
    if thinking is not None:
        kwargs["thinking_config"] = thinking
    return types.GenerateContentConfig(**kwargs)


def _is_thinking_unsupported(exc):
    """True when a request failed because the model rejected thinking_level."""
    message = str(exc).lower()
    return "thinking" in message and any(
        marker in message
        for marker in ("not supported", "unsupported", "unknown", "invalid", "cannot")
    )


def _finish_reason(response):
    try:
        candidate = (response.candidates or [None])[0]
        reason = getattr(candidate, "finish_reason", None)
        return str(getattr(reason, "name", reason) or "")
    except Exception:
        return ""


def _extract_text(response):
    """Collect the visible text from a Gemini response.

    ``response.text`` can be ``None`` (safety block, tool call, or a
    thinking-only turn), so the candidate parts are walked as a fallback and
    "thought" parts are skipped.
    """
    try:
        # The `text` accessor can raise on responses with non-text parts.
        text = response.text
    except Exception:
        text = None
    if text:
        return text

    chunks = []
    for candidate in getattr(response, "candidates", None) or []:
        content = getattr(candidate, "content", None)
        for part in getattr(content, "parts", None) or []:
            if getattr(part, "thought", False):
                continue
            part_text = getattr(part, "text", None)
            if part_text:
                chunks.append(part_text)
    return "".join(chunks)


def _blocked_message(response):
    """Return a human-readable reason when a response carried no text."""
    feedback = getattr(response, "prompt_feedback", None)
    blocked = getattr(feedback, "block_reason", None)
    if blocked:
        name = getattr(blocked, "name", blocked)
        detail = getattr(feedback, "block_reason_message", None)
        return f"prompt blocked by Gemini safety filters ({name}){': ' + detail if detail else ''}"
    reason = _finish_reason(response)
    if reason and reason not in ("STOP", "FINISH_REASON_UNSPECIFIED"):
        return f"response stopped early (finish_reason={reason})"
    return "empty response from Gemini"


_JSON_FENCE_RE = re.compile(r"^\s*```(?:json|JSON)?\s*(.*?)\s*```\s*$", re.DOTALL)


def coerce_json_text(text):
    """Normalise Gemini output so ``json.loads`` behaves like OpenAI JSON mode.

    Gemini honours ``response_mime_type="application/json"`` but can still wrap
    the payload in a markdown fence or add a short preamble, especially on
    lite/flash models. This strips fences and, failing that, slices out the
    outermost JSON object/array. Text that is already valid JSON is returned
    unchanged; anything unparseable is returned as-is so the caller's existing
    JSONDecodeError fallback still applies.
    """
    if not isinstance(text, str):
        return text
    candidate = text.strip()
    if not candidate:
        return text

    try:
        json.loads(candidate)
        return candidate
    except ValueError:
        pass

    fenced = _JSON_FENCE_RE.match(candidate)
    if fenced:
        inner = fenced.group(1).strip()
        try:
            json.loads(inner)
            return inner
        except ValueError:
            candidate = inner

    for opener, closer in (("{", "}"), ("[", "]")):
        start = candidate.find(opener)
        end = candidate.rfind(closer)
        if start != -1 and end > start:
            sliced = candidate[start:end + 1]
            try:
                json.loads(sliced)
                return sliced
            except ValueError:
                continue
    return text


def generate_text(prompt, model=None, response_format=None, temperature=0.5, seed=None):
    """Generate content with Gemini and return plain text.

    Mirrors the OpenAI chat-completions call used elsewhere in the app:
    a single user message in, message content out.

    Raises:
        RuntimeError: on configuration problems or an unusable response.
    """
    from google.genai import types

    client = get_client()
    model_name = get_model_name(model)
    max_tokens = _max_output_tokens()

    def _generate(max_output_tokens, thinking):
        return client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=_build_config(
                types, response_format, temperature, seed, max_output_tokens, thinking
            ),
        )

    try:
        response = _generate(max_tokens, _thinking_config(types))
    except Exception as exc:
        # Pre-Gemini-3 models reject thinking_level; fall back to defaults.
        if not _is_thinking_unsupported(exc):
            raise
        response = _generate(max_tokens, None)

    content = _extract_text(response)

    # A thinking model can spend the whole output budget before emitting any
    # visible text. Retry once with a larger budget rather than failing.
    if not content and _finish_reason(response) == "MAX_TOKENS":
        response = _generate(max_tokens * 2, None)
        content = _extract_text(response)

    if not content:
        raise RuntimeError(_blocked_message(response))

    if response_format == "json_object":
        content = coerce_json_text(content)
    return content
