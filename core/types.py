from dataclasses import dataclass
from typing import List, Dict

@dataclass
class ExtractOptions:
    exclude_selectors: list[str]
    annotate_links: bool = False
    remove_before_h1: bool = False
    include_img_src: bool = False

Meta = Dict[str, str]  # or create a dataclass if you prefer strict typing
