# --- Constants ---
# Optimized stop words for food products - removed some that might be important for food
STOP_WORDS = set([
    "a", "an", "the", "and", "or", "but", "in", "on", "at", "to",
    "of", "with", "by", "is", "am", "are", "was", "were", "be", "been",
    "being", "have", "has", "had", "do", "does", "did", "case", "pack",
    # Food-specific stop words that don't add matching value
    "brand", "product", "item", "food", "natural", "premium", "quality"
])

UNIT_CONVERSION_MAP = {
    'oz': 28.35, 'lb': 453.592, 'gallon': 3785.41, 'fl oz': 29.5735,
    'l': 1000, 'ml': 1, 'g': 1, 'kg': 1000
}
