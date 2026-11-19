"""Chain name normalization utility for matching chain names across different data sources.

Normalizes chain names by:
- Trimming whitespace
- Lowercasing
- Removing common variations (quotes, extra spaces)
- Allowing fuzzy matching for known aliases
"""

import re
from typing import Optional, Dict, Set
from pathlib import Path
import json


# Known chain aliases - maps variations to canonical name
CHAIN_ALIASES = {
    'аптека низькі ціни': 'Низькі Ціни',
    'низкі ціни': 'Низькі Ціни',
    'аптеки низькі ціни': 'Низькі Ціни',
    'аптека "низкі ціни"': 'Низькі Ціни',
    'аптека\'низкі ціни\'': 'Низькі Ціни',
    # Add more aliases as needed
}

# Canonical chain names (proper capitalization)
CANONICAL_CHAINS = {
    'Низькі Ціни',
    'Аптека 9-1-1',
    'Аптеки Запоріжжя',
    'Товари для здоров\'я',
    # Add more as discovered
}


def normalize_chain_name(name: str) -> str:
    """Normalize a chain name for comparison.

    Args:
        name: Raw chain name from any source

    Returns:
        Normalized lowercase name for comparison
    """
    if not name:
        return ''

    # Convert to string, strip whitespace
    s = str(name).strip()

    # Remove quotes
    s = s.strip('\'"')

    # Normalize whitespace (multiple spaces -> single)
    s = re.sub(r'\s+', ' ', s)

    # Lowercase for comparison
    return s.lower()


def get_canonical_chain_name(name: str) -> str:
    """Get canonical (properly capitalized) chain name.

    Args:
        name: Raw or normalized chain name

    Returns:
        Canonical chain name with proper capitalization
    """
    normalized = normalize_chain_name(name)

    # Check aliases first
    if normalized in CHAIN_ALIASES:
        return CHAIN_ALIASES[normalized]

    # Check if it matches a known canonical name (case-insensitive)
    for canonical in CANONICAL_CHAINS:
        if normalize_chain_name(canonical) == normalized:
            return canonical

    # Fallback: title case the normalized name
    return ' '.join(word.capitalize() for word in normalized.split())


def build_chain_mapping(df_req, df_pharmacy=None, df_mhl=None, df_offtake=None) -> Dict[str, str]:
    """Build a mapping from all known chain name variations to canonical names.

    Args:
        df_req: Requisites DataFrame with 'Клиент из базы данных' or similar column
        df_pharmacy: Pharmacy DataFrame with CHAIN_NAME column
        df_mhl: MHL DataFrame with CHAIN_NAME column
        df_offtake: Offtake DataFrame with CHAIN_NAME column

    Returns:
        Dict mapping normalized name -> canonical name
    """
    mapping = {}

    # Collect all chain names from all sources
    all_names = set()

    # From requisites
    if df_req is not None and not df_req.empty:
        for col in ['Клиент из базы данных', 'Клиент с базы данных', 'Мережа', 'Назва', 'chain_name']:
            if col in df_req.columns:
                names = df_req[col].dropna().astype(str).str.strip()
                all_names.update(names[names != ''].tolist())

    # From pharmacy data
    if df_pharmacy is not None and not df_pharmacy.empty and 'CHAIN_NAME' in df_pharmacy.columns:
        names = df_pharmacy['CHAIN_NAME'].dropna().astype(str).str.strip()
        all_names.update(names[names != ''].tolist())

    # From MHL data
    if df_mhl is not None and not df_mhl.empty and 'CHAIN_NAME' in df_mhl.columns:
        names = df_mhl['CHAIN_NAME'].dropna().astype(str).str.strip()
        all_names.update(names[names != ''].tolist())

    # From offtake data
    if df_offtake is not None and not df_offtake.empty and 'CHAIN_NAME' in df_offtake.columns:
        names = df_offtake['CHAIN_NAME'].dropna().astype(str).str.strip()
        all_names.update(names[names != ''].tolist())

    # Build mapping: normalized -> canonical
    for name in all_names:
        normalized = normalize_chain_name(name)
        canonical = get_canonical_chain_name(name)
        mapping[normalized] = canonical

    return mapping


def find_matching_chain(target: str, candidates: Set[str]) -> Optional[str]:
    """Find matching chain from candidates using normalized comparison.

    Args:
        target: Target chain name (e.g., from UI selection)
        candidates: Set of candidate chain names (e.g., from requisites)

    Returns:
        Matching canonical candidate name or None
    """
    target_norm = normalize_chain_name(target)

    for candidate in candidates:
        if normalize_chain_name(candidate) == target_norm:
            return candidate

    return None


def load_chain_aliases() -> Dict[str, str]:
    """Load chain aliases from JSON config file."""
    alias_file = Path(__file__).resolve().parent / 'chain_aliases.json'
    if alias_file.exists():
        try:
            with open(alias_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_chain_aliases(aliases: Dict[str, str]) -> bool:
    """Save chain aliases to JSON config file."""
    alias_file = Path(__file__).resolve().parent / 'chain_aliases.json'
    try:
        with open(alias_file, 'w', encoding='utf-8') as f:
            json.dump(aliases, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False