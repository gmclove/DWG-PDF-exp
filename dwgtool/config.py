
# Configuration for title block scanning and formatting
from typing import Dict, Any, Set, Tuple

# 1. Default block names for sheet-specific title block (UPPERCASE)
DEFAULT_TARGET_BLOCK_NAMES: Set[str] = {
    # Example: update to your actual names
    "GF MALTA TITLE BLOCK 30X42-TB-ATT",
}

# 2. Sheet number components (we match by TagString OR PromptString)
DEFAULT_SHEET_TOP_TAGS: Tuple[str, ...] = ("FC-E", "TOPNUMBER", "TOP_NUM", "TOP")
DEFAULT_SHEET_TOP_PROMPTS: Tuple[str, ...] = ("Top Number",)

DEFAULT_SHEET_BOTTOM_TAGS: Tuple[str, ...] = ("442C", "BOTTOMNUMBER", "BOTTOM_NUM", "BOTTOM")
DEFAULT_SHEET_BOTTOM_PROMPTS: Tuple[str, ...] = ("Bottom Number",)

# 3. Title lines (TITLE_1 through TITLE_5) + alias seen in your sample
DEFAULT_TITLE_TAG_PRIMARY: Tuple[str, ...] = ("TITLE_1", "TITLE_2", "TITLE_3", "TITLE_4", "TITLE_5")
DEFAULT_TITLE_PROMPT_PRIMARY: Tuple[str, ...] = ("TITLE_1", "TITLE_2", "TITLE_3", "TITLE_4", "TITLE_5")
DEFAULT_TITLE_PROMPT_ALIAS: Tuple[str, ...] = ("ELECTRICAL",)

# 4. Revision index range R0..R9
DEFAULT_REVISION_INDEX_RANGE = range(0, 10)

# 5. Sheet number separator and title joiner
DEFAULT_SHEETNO_SEPARATOR = "-"
DEFAULT_TITLE_JOINER = " "

def prompt_for_titleblock_config() -> Dict[str, Any]:
    print("Title Block Configuration:")
    use_default = input("Use default title block settings? (Y/N): ").strip().lower()

    if use_default.startswith('y'):
        return {
            "target_block_names": DEFAULT_TARGET_BLOCK_NAMES,
            "sheet_top_tags": DEFAULT_SHEET_TOP_TAGS,
            "sheet_top_prompts": DEFAULT_SHEET_TOP_PROMPTS,
            "sheet_bottom_tags": DEFAULT_SHEET_BOTTOM_TAGS,
            "sheet_bottom_prompts": DEFAULT_SHEET_BOTTOM_PROMPTS,
            "title_tag_primary": DEFAULT_TITLE_TAG_PRIMARY,
            "title_prompt_primary": DEFAULT_TITLE_PROMPT_PRIMARY,
            "title_prompt_alias": DEFAULT_TITLE_PROMPT_ALIAS,
            "revision_index_range": DEFAULT_REVISION_INDEX_RANGE,
            "sheetno_separator": DEFAULT_SHEETNO_SEPARATOR,
            "title_joiner": DEFAULT_TITLE_JOINER,
        }

    print("Enter override configuration (press ENTER to keep default for each):")

    def ask_list(prompt: str, default_tuple: Tuple[str, ...]) -> Tuple[str, ...]:
        raw = input(f"{prompt} [{', '.join(default_tuple)}]: ").strip()
        if not raw:
            return default_tuple
        return tuple(x.strip() for x in raw.split(','))

    raw_blocks = input(
        f"Sheet title block names (comma-separated) [{', '.join(DEFAULT_TARGET_BLOCK_NAMES) if DEFAULT_TARGET_BLOCK_NAMES else 'NONE'}]: "
    ).strip()
    target_block_names = (
        {s.strip().upper() for s in raw_blocks.split(',')} if raw_blocks else DEFAULT_TARGET_BLOCK_NAMES
    )

    cfg = {
        "target_block_names": target_block_names,
        "sheet_top_tags": ask_list("Sheet top number tags", DEFAULT_SHEET_TOP_TAGS),
        "sheet_top_prompts": ask_list("Sheet top number prompts", DEFAULT_SHEET_TOP_PROMPTS),
        "sheet_bottom_tags": ask_list("Sheet bottom number tags", DEFAULT_SHEET_BOTTOM_TAGS),
        "sheet_bottom_prompts": ask_list("Sheet bottom number prompts", DEFAULT_SHEET_BOTTOM_PROMPTS),
        "title_tag_primary": ask_list("Title line tag names", DEFAULT_TITLE_TAG_PRIMARY),
        "title_prompt_primary": ask_list("Title line prompts", DEFAULT_TITLE_PROMPT_PRIMARY),
        "title_prompt_alias": ask_list("Title alias prompts", DEFAULT_TITLE_PROMPT_ALIAS),
        "revision_index_range": DEFAULT_REVISION_INDEX_RANGE,
        "sheetno_separator": input(
            f"Sheet number separator [{DEFAULT_SHEETNO_SEPARATOR}]: "
        ).strip() or DEFAULT_SHEETNO_SEPARATOR,
        "title_joiner": input(
            f"Sheet title joiner [{DEFAULT_TITLE_JOINER}]: "
        ).strip() or DEFAULT_TITLE_JOINER,
    }
    return cfg
