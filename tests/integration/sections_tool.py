"""
title: Sections
"""


class Tools:
    async def read_section(self, number: int) -> str:
        """
        Reads a section of the test document.
        :param number: Number of the section
        """
        base = f"Section {number}: content of the tool result that fills the context. "
        return (base * 60)[:3500]
