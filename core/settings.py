# -------------------------
# CONFIG / CONSTANTS
# -------------------------

ALWAYS_STRIP = {"script", "style", "noscript", "template"}

# Inline tags to treat as part of the same paragraph when grouping text
INLINE_TAGS = {"a","span","strong","em","b","i","u","s","small","sup","sub","mark","abbr","time","code","var","kbd"}

DEFAULT_EXCLUDE = [
    "header", "footer", "nav",
    ".cookie", ".newsletter",
    "[class*='breadcrumb']",
    "[class*='wishlist']",
    "[class*='simplesearch']",
    "[id*='gallery']",
    "[class*='usp']",
    "[class*='feefo']",
    "[class*='associated-blogs']",
    "[class*='popular']",
    # Explore/SPA results containers and variants
    ".sr-main.js-searchpage-content.visible",
    "[class~='sr-main'][class~='js-searchpage-content'][class~='visible']",
    "[class*='js-searchpage-content']",
    "[class*='searchpage-content']",
    # Map modal container to exclude
    ".lmd-map-modal-create.js-lmd-map-modal-map",
]

DATE_TZ = "Europe/London"
DATE_FMT = "%d/%m/%Y"  # UK format

# Common UI/analytics noise to drop when emitting <p>
NOISE_SUBSTRINGS = (
    "google tag manager",
    "loading results",
    "load more",
    "updating results",
    "something went wrong",
    "filters",
    "apply filters",
    "clear",
    "sort by",
    "to collect end-user usage analytics",
    "place this code immediately before the closing",
)

SIDEBAR_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap');

/* Global font */
html, body, [data-testid="stAppViewContainer"] *:not(.material-icons):not(.material-icons-outlined):not(.material-symbols-outlined):not(.material-symbols-rounded):not(.material-symbols-sharp) {
  font-family: 'Montserrat', sans-serif;
}
/* Restore icon fonts so ligatures render as icons */
.material-icons,
.material-icons-outlined,
.material-symbols-outlined,
.material-symbols-rounded,
.material-symbols-sharp {
  font-family: 'Material Icons','Material Icons Outlined','Material Symbols Outlined','Material Symbols Rounded','Material Symbols Sharp' !important;
  font-weight: normal;
  font-style: normal;
  line-height: 1;
  -webkit-font-feature-settings: 'liga';
  -webkit-font-smoothing: antialiased;
}

/* Main title: target first H1 robustly */
section[tabindex="0"] h1:first-of-type {
  text-align: center;
  color: #4A90E2;
  font-size: 3em;
  padding-bottom: .5em;
  border-bottom: 2px solid #4A90E2;
}

/* Sidebar look + width */
[data-testid="stSidebar"] {
  background-color: #1a1e24;
  border-right: 1px solid #4A90E2;
  min-width: 320px;
  max-width: 420px;
}

/* Expander headers */
[data-testid="stExpander"] [data-testid="stExpanderHeader"] {
  background-color: #363945;
  border-radius: 8px;
  padding: 10px 15px;
  margin-bottom: 10px;
  border: none;
  font-weight: bold;
  color: #E0E0E0;
}

/* Buttons */
.stButton > button {
  width: 100%;
  background-color: #323640;
  color: #E0E0E0;
  border: 1px solid #4A90E2;
  border-radius: 8px;
  padding: 10px;
  transition: background-color .3s, color .3s;
}
.stButton > button:hover {
  background-color: #4A90E2;
  color: #fff;
  border-color: #fff;
}

/* Tabs */
[data-testid="stTabs"] button[role="tab"] { background-color: #323640; color: #E0E0E0; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
  color: #4A90E2;
  box-shadow: inset 0 -3px 0 0 #4A90E2;
}
</style>
"""
