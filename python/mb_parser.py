"""
Mount & Blade Warband TXT file parser - Python proof of concept.

Ported from MnBEditor.Core.FileParser (C#) / FileParser.bas (VB6).
Demonstrates the same parsing logic in Python for the MOD community.

Usage:
    from mb_parser import FileParser, ItemLoader

    repo = ItemLoader.load_items(FileParser("item_kinds1.txt"))
    print(f"Loaded {len(repo['items'])} items")
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional
import struct


class FileParser:
    """Low-level tokenizer for M&B space-delimited TXT files."""

    def __init__(self, source):
        if isinstance(source, str):
            with open(source, 'rb') as f:
                self._buffer = f.read()
        elif isinstance(source, (bytes, bytearray)):
            self._buffer = bytes(source)
        else:
            raise TypeError(f"Expected str (filepath), bytes, or bytearray, got {type(source)}")
        self._position = 0
        self._length = len(self._buffer)

    @property
    def is_eof(self) -> bool:
        return self._position >= self._length

    def get_word(self) -> str:
        """Reads the next whitespace-delimited token."""
        chars = []
        while self._position < self._length:
            b = self._buffer[self._position]
            self._position += 1
            if b in (10, 13, 32):  # LF, CR, space
                if chars:
                    break
            else:
                chars.append(chr(b))
        return ''.join(chars)

    def get_long(self) -> int:
        """Reads the next token as an integer."""
        word = self.get_word()
        try:
            return int(word)
        except ValueError:
            return 0

    def get_double(self) -> float:
        """Reads the next token as a float."""
        word = self.get_word()
        try:
            return float(word)
        except ValueError:
            return 0.0

    def get_line(self) -> str:
        """Reads until CR/LF, consuming the line ending."""
        chars = []
        while self._position < self._length:
            b = self._buffer[self._position]
            self._position += 1
            if b == 13:  # CR
                if self._position < self._length and self._buffer[self._position] == 10:
                    self._position += 1
                break
            elif b == 10:  # LF
                break
            else:
                chars.append(chr(b))
        return ''.join(chars)

    def skip_words(self, count: int) -> None:
        for _ in range(count):
            self.get_word()


# ---- Data Models ----

@dataclass
class Item:
    id: int = 0
    db_name: str = ""
    display_name: str = ""
    texture_name: str = ""
    mesh_count: int = 0
    mesh_names: List[str] = field(default_factory=list)
    mesh_params: List[str] = field(default_factory=list)
    item_type: str = ""
    action: str = ""
    price: int = 0
    prefix: str = ""
    weight: str = ""
    abundance: int = 0
    head_armor: int = 0
    body_armor: int = 0
    leg_armor: int = 0
    difficulty: int = 0
    hit_points: int = 0
    speed_rating: int = 0
    missile_speed: int = 0
    weapon_length: int = 0
    max_ammo: int = 0
    thrust_damage: int = 0
    swing_damage: int = 0
    faction_count: int = 0
    factions: List[int] = field(default_factory=list)
    trigger_count: int = 0
    csv_name: str = ""


@dataclass
class Troop:
    id: int = 0
    str_id: str = ""
    name: str = ""
    plural_name: str = ""
    unknown_warband: str = ""
    flags: str = ""
    scene: int = 0
    reserved: int = 0
    faction: int = 0
    upgrade1: int = 0
    upgrade2: int = 0
    inventory: List[tuple] = field(default_factory=lambda: [(0, 0)] * 64)
    attributes: Dict[str, int] = field(default_factory=lambda: {
        'str': 0, 'agi': 0, 'int': 0, 'cha': 0, 'level': 0
    })
    proficiencies: Dict[str, int] = field(default_factory=lambda: {
        'one_handed': 0, 'two_handed': 0, 'polearm': 0,
        'archery': 0, 'crossbow': 0, 'throwing': 0, 'firearm': 0
    })
    skills: List[int] = field(default_factory=lambda: [0] * 6)
    face: List[str] = field(default_factory=lambda: [""] * 8)
    csv_name: str = ""


@dataclass
class DataRepository:
    items: List[Item] = field(default_factory=list)
    troops: List[Troop] = field(default_factory=list)
    factions: List[dict] = field(default_factory=list)
    language: str = "en"


# ---- Item Loader ----

class ItemLoader:
    """Loads and saves item_kinds1.txt files."""

    @staticmethod
    def load_items(repo: DataRepository, parser: FileParser) -> None:
        # Version header
        parser.get_word()  # "itemsfile"
        parser.get_word()  # "version"
        parser.get_word()  # "3"
        count = parser.get_long()

        repo.items.clear()
        for i in range(count):
            item = ItemLoader._read_item(parser)
            item.id = i
            repo.items.append(item)

    @staticmethod
    def _read_item(parser: FileParser) -> Item:
        item = Item()
        item.db_name = parser.get_word()
        item.display_name = parser.get_word()
        item.texture_name = parser.get_word()

        item.mesh_count = parser.get_long()
        for _ in range(item.mesh_count):
            item.mesh_names.append(parser.get_word())
            item.mesh_params.append(parser.get_word())

        item.item_type = parser.get_word()
        item.action = parser.get_word()
        item.price = parser.get_long()
        item.prefix = parser.get_word()
        item.weight = parser.get_word()
        item.abundance = parser.get_long()
        item.head_armor = parser.get_long()
        item.body_armor = parser.get_long()
        item.leg_armor = parser.get_long()
        item.difficulty = parser.get_long()
        item.hit_points = parser.get_long()
        item.speed_rating = parser.get_long()
        item.missile_speed = parser.get_long()
        item.weapon_length = parser.get_long()
        item.max_ammo = parser.get_long()
        item.thrust_damage = parser.get_long()
        item.swing_damage = parser.get_long()

        item.faction_count = parser.get_long()
        for _ in range(item.faction_count):
            item.factions.append(parser.get_long())

        item.trigger_count = parser.get_long()
        for _ in range(item.trigger_count):
            parser.get_double()  # trigger type
            act_count = parser.get_long()
            for _ in range(act_count):
                parser.get_word()  # opcode
                param_count = parser.get_long()
                parser.skip_words(param_count)

        return item


# ---- Tests (run with: python mb_parser.py) ----

if __name__ == "__main__":
    # Test 1: basic parsing
    print("Test 1: FileParser basic tokenization...")
    p = FileParser(b"hello world 123\r\ntoken2")
    assert p.get_word() == "hello"
    assert p.get_word() == "world"
    assert p.get_long() == 123
    assert p.get_word() == "token2"
    print("  PASS")

    # Test 2: item parsing
    print("Test 2: Item parsing...")
    data = (
        "itemsfile version 3 2\r\n"
        " itm_sword Sword sword_tex 1 mesh1 0 1 655370 1000 0 2.5 100 "
        "0 0 0 0 100 90 0 95 0 10 30 0 0\r\n"
        " itm_shield Shield shield_tex 1 mesh2 0 7 167772194 500 0 3.0 50 "
        "0 0 0 0 80 0 0 0 0 0 0 0 0\r\n"
    )
    repo = DataRepository()
    ItemLoader.load_items(repo, FileParser(data.encode()))

    assert len(repo.items) == 2
    assert repo.items[0].db_name == "itm_sword"
    assert repo.items[0].price == 1000
    assert repo.items[0].weight == "2.5"
    assert repo.items[1].db_name == "itm_shield"
    assert repo.items[1].price == 500
    print("  PASS")

    # Test 3: empty file
    print("Test 3: Empty file handling...")
    p = FileParser(b"")
    assert p.get_word() == ""
    assert p.is_eof
    print("  PASS")

    # Test 4: CSV parsing (Dictionary O(n) pattern)
    print("Test 4: Dictionary-based CSV matching...")
    items_dict = {"itm_sword": Item(db_name="itm_sword"), "itm_shield": Item(db_name="itm_shield")}
    csv_data = "itm_sword|Iron Sword\nitm_shield|Wooden Shield\n"
    for line in csv_data.strip().split("\n"):
        key, value = line.split("|")
        key_lower = key.lower()
        if key_lower in items_dict:
            items_dict[key_lower].csv_name = value
    assert items_dict["itm_sword"].csv_name == "Iron Sword"
    assert items_dict["itm_shield"].csv_name == "Wooden Shield"
    print("  PASS")

    print(f"\nAll 4 tests passed!")
