"""序列處理相關工具函式

Sequence processing utility functions.
"""


def reverse_complement(seq: str) -> str:
    """計算反向互補序列

    Calculate the reverse complement of a DNA sequence.

    Args:
        seq (str): DNA 序列 / DNA sequence.

    Returns:
        str: 反向互補序列 / Reverse complement sequence.
    """
    complement = {"A": "T", "T": "A", "C": "G", "G": "C"}
    return "".join(complement.get(base, base) for base in reversed(seq))
