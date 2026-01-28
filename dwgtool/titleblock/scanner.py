
import re
from typing import Dict, Any, Tuple, List

def _norm_tag(s: str) -> str:
    return re.sub(r'[^A-Z0-9]', '', (s or '').upper())

def _norm_prompt(s: str) -> str:
    return re.sub(r'[^A-Z0-9]', '', (s or '').upper())

def _norm_name(s: str) -> str:
    return (s or '').upper().strip()

def read_titleblock_from_active_layout_robust(
    doc,
    com_get,
    com_call,
    cfg: Dict[str, Any],
) -> Tuple[str, str, str, str, str, str, Dict[str, str]]:
    """
    Read the sheet-specific title block attributes on the ACTIVE layout (PaperSpace).
    Returns 7 values:
      (sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs_dict)

    - Uses fast block-name filter if cfg['target_block_names'] provided.
    - Otherwise scans & scores blocks, matching by TagString AND PromptString.
    - Skips XREF blocks where detectable.
    """
    target_block_names = cfg.get('target_block_names') or set()
    sheet_top_tags = tuple(cfg.get('sheet_top_tags') or ())
    sheet_top_prompts = tuple(cfg.get('sheet_top_prompts') or ())
    sheet_bottom_tags = tuple(cfg.get('sheet_bottom_tags') or ())
    sheet_bottom_prompts = tuple(cfg.get('sheet_bottom_prompts') or ())
    title_tag_primary = tuple(cfg.get('title_tag_primary') or ())
    title_prompt_primary = tuple(cfg.get('title_prompt_primary') or ())
    title_prompt_alias = tuple(cfg.get('title_prompt_alias') or ())
    revision_index_range = cfg.get('revision_index_range') or range(0, 10)
    sheetno_sep = cfg.get('sheetno_separator') or '-'
    title_joiner = cfg.get('title_joiner') or ' '

    tb_names = {_norm_name(n) for n in target_block_names} if target_block_names else None
    sheet_top_tag_norm = {_norm_tag(t) for t in sheet_top_tags}
    sheet_top_prompt_norm = {_norm_prompt(p) for p in sheet_top_prompts}
    sheet_bot_tag_norm = {_norm_tag(t) for t in sheet_bottom_tags}
    sheet_bot_prompt_norm = {_norm_prompt(p) for p in sheet_bottom_prompts}
    title_tag_norm_order = [_norm_tag(t) for t in title_tag_primary]
    title_prompt_norm_order = [_norm_prompt(p) for p in title_prompt_primary]
    title_prompt_alias_norm = {_norm_prompt(p) for p in title_prompt_alias}

    sheet_no = 'Not Found'
    sheet_title = 'Not Found'
    rev_no = 'Not Found'
    rev_date = 'Not Found'
    rev_desc = 'Not Found'
    rev_by = 'Not Found'
    other_attrs: Dict[str, str] = {}

    try:
        ps = com_get(doc, 'PaperSpace')
        ps_count = com_get(ps, 'Count')
    except Exception:
        return sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs

    def collect_attrs(ent) -> List[Dict[str, str]]:
        out: List[Dict[str, str]] = []
        try:
            if not com_get(ent, 'HasAttributes'):
                return out
            arr = com_call(ent, 'GetAttributes')
        except Exception:
            return out
        for a in arr:
            try:
                raw_tag = com_get(a, 'TagString') or ''
            except Exception:
                raw_tag = ''
            try:
                raw_prompt = com_get(a, 'PromptString') or ''
            except Exception:
                raw_prompt = ''
            norm_tag = _norm_tag(raw_tag)
            norm_prompt = _norm_prompt(raw_prompt)
            try:
                val = (com_get(a, 'TextString') or '').strip()
            except Exception:
                val = ''
            out.append({
                'raw_tag': raw_tag, 'norm_tag': norm_tag,
                'raw_prompt': raw_prompt, 'norm_prompt': norm_prompt,
                'value': val,
            })
        return out

    def analyze_block(attrs: List[Dict[str, str]]):
        score = 0
        top = ''
        bot = ''
        titles = {t: '' for t in title_tag_norm_order}
        revs = {}

        # Sheet number + titles
        for a in attrs:
            val = a['value']
            if not val:
                continue
            nt = a['norm_tag']
            np = a['norm_prompt']
            if (nt in sheet_top_tag_norm or np in sheet_top_prompt_norm) and not top:
                top = val; score += 3; continue
            if (nt in sheet_bot_tag_norm or np in sheet_bot_prompt_norm) and not bot:
                bot = val; score += 3; continue
            if nt in titles and not titles[nt]:
                titles[nt] = val; score += 1; continue
            if np in title_prompt_norm_order:
                idx = title_prompt_norm_order.index(np)
                tnorm = title_tag_norm_order[idx]
                if not titles[tnorm]:
                    titles[tnorm] = val; score += 1; continue

        # Alias to TITLE_1
        if title_tag_norm_order:
            t1 = title_tag_norm_order[0]
            if not titles[t1]:
                for a in attrs:
                    if not a['value']:
                        continue
                    if a['norm_prompt'] in title_prompt_alias_norm or a['norm_tag'] == _norm_tag('ELECTRICAL'):
                        titles[t1] = a['value']; score += 1; break

        # Revisions R#NO/DATE/DESC/BY
        for a in attrs:
            val = a['value']
            if not val:
                continue
            m = re.match(r'^R(\d+)(NO|DATE|DESC|BY)$', a['norm_tag'])
            if m:
                idx = int(m.group(1))
                if idx in revision_index_range:
                    kind = m.group(2)
                    if idx not in revs:
                        revs[idx] = { 'NO':'', 'DATE':'', 'DESC':'', 'BY':'' }
                    revs[idx][kind] = val
                    score += 1

        return score, top, bot, titles, revs

    best = {
        'score': -1, 'attrs': [], 'top': '', 'bot': '',
        'titles': {t: '' for t in title_tag_norm_order}, 'revs': {},
        'attrs_count': 0, 'blkname': ''
    }

    def consider_entity(ent):
        nonlocal best
        try:
            blkname = com_get(ent, 'Name') or ''
        except Exception:
            blkname = ''
        blkname_norm = _norm_name(blkname)

        # Skip XREFs
        try:
            if com_get(ent, 'IsXRef'):
                return
        except Exception:
            pass

        # Name filter if provided
        if tb_names is not None and blkname_norm not in tb_names:
            return

        attrs = collect_attrs(ent)
        if not attrs:
            return
        score, top, bot, titles, revs = analyze_block(attrs)
        if (score > best['score']) or (score == best['score'] and len(attrs) > best['attrs_count']):
            best.update({
                'score': score, 'attrs': attrs, 'top': top, 'bot': bot,
                'titles': titles, 'revs': revs, 'attrs_count': len(attrs),
                'blkname': blkname,
            })

    for i in range(ps_count):
        try:
            ent = com_call(ps, 'Item', i)
        except Exception:
            continue
        consider_entity(ent)

    if best['score'] < 0:
        return sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs

    # Compose outputs
    if best['top'] and best['bot']:
        sheet_no = f"{best['top']}{sheetno_sep}{best['bot']}"
    elif best['top']:
        sheet_no = best['top']
    elif best['bot']:
        sheet_no = best['bot']

    title_vals = [best['titles'][t] for t in title_tag_norm_order if best['titles'][t]]
    if title_vals:
        sheet_title = title_joiner.join(title_vals)

    if best['revs']:
        max_idx = max((idx for idx, d in best['revs'].items() if d.get('NO')), default=None)
        if max_idx is not None:
            d = best['revs'][max_idx]
            rev_no = d.get('NO') or rev_no
            rev_date = d.get('DATE') or rev_date
            rev_desc = d.get('DESC') or rev_desc
            rev_by = d.get('BY') or rev_by

    # Other attributes (raw tag names as keys), excluding used fields
    used_norm_tags = set(sheet_top_tag_norm) | set(sheet_bot_tag_norm)
    used_norm_prompts = set(sheet_top_prompt_norm) | set(sheet_bot_prompt_norm)
    used_norm_tags |= set(title_tag_norm_order)
    used_norm_prompts |= set(title_prompt_norm_order) | set(title_prompt_alias_norm)
    for idx in revision_index_range:
        for kind in ('NO','DATE','DESC','BY'):
            used_norm_tags.add(_norm_tag(f'R{idx}{kind}'))

    for a in best['attrs']:
        if not a['value']:
            continue
        if a['norm_tag'] in used_norm_tags or a['norm_prompt'] in used_norm_prompts:
            continue
        key = a['raw_tag'] or a['raw_prompt'] or 'ATTR'
        if key not in other_attrs:
            other_attrs[key] = a['value']

    return sheet_no, sheet_title, rev_no, rev_date, rev_desc, rev_by, other_attrs
