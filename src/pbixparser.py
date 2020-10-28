"""
@author: Yannick ARMENTI

"""
import tempfile
import zipfile
from pathlib import Path
import json
import codecs
import copy
from typing import Any


class PBIXExtractor(object):
    """
    This class allows to read and write a PBIX file from disk.

    Attributes:
        path (pathlib.Path): PBIX filepath on disk.
    """

    def __init__(self, pbix_file: str) -> None:
        """
        Constructor.

        Args:
            pbix_file (str): PBIX file path as string.
        """
        self.path = Path(pbix_file)
        self.__folder_path = None

    def extract_pbix(self) -> Path:
        """
        Extracts the pbix file using zipfile module.
        The extraction is performed in a temporary drive on the disk.

        Returns:
            pathlib.Path: Output extraction folder of the PBIX file.
        """
        self.__folder_path = tempfile.TemporaryDirectory()
        # self.__folder_path = Path(self.path.stem)

        with zipfile.ZipFile(self.path, "r") as zip_ref:
            zip_ref.extractall(self.__folder_path.name)

        return Path(self.__folder_path.name)

    def export_pbix(self, output_path: str) -> None:
        """
        Exports the working directory to a PBIX file.

        Args:
            output_path (str): Path of the PBIX output file.
        """

        # Remove SecurityBindings
        folder_path = Path(self.__folder_path.name)
        (folder_path / "SecurityBindings").unlink()

        length = len(folder_path.parts)

        with zipfile.ZipFile(output_path, "w") as zip_ref:
            for i in folder_path.glob('**/*'):
                zip_ref.write(i, Path(*i.parts[length:]))

        self.__folder_path.cleanup()


class PBIXManager(object):
    """
    Base class allowing encapsulation of PBIX modification.

    Attributes:
        extract_path (Path): Output extraction folder of the PBIX file.
        extractor (PBIXExtractor): PBIX Extractor object.
    """

    def __init__(self, pbix_file: str) -> None:
        """
        Constructor.

        Args:
            pbix_file (str): PBIX file path as string.
        """
        self.extractor = PBIXExtractor(pbix_file)
        self.extract_path = None

    def extract(self) -> 'PBIXManager':
        """
        Extracts the PBIX file to disk and sets the `extract_path` instance
        variable.

        Returns:
            PBIXManager: self.
        """
        self.extract_path = self.extractor.extract_pbix()
        return self

    def save(self, out_pbix: str) -> None:
        """
        Exports the PBIX to disk.

        Args:
            out_pbix (str): Path of the PBIX output file.
        """
        self.extractor.export_pbix(out_pbix)


class PBIXSectionManager(PBIXManager):
    """
    PBIX Section modifier (adding, modifying)

    Attributes:
        encoding (str): Encoding of the Layout file.
        layout (dict): Dictionary with all the PBIX Layout content.
        special_parsing_keys (list): Keys that should be parsed separately.
    """

    encoding = "utf-16-LE"

    special_parsing_keys = [
        "config",
        "filters"
    ]

    def __init__(self, pbix_file: str) -> None:
        """
        Constructor.

        Args:
            pbix_file (str): PBIX file path as string.
        """
        super().__init__(pbix_file)
        self.layout = None

    def __recursive_parser_str2json(self, data: Any) -> dict:
        """
        Recursive parser from string to JSON dict object.

        Args:
            data (Any): Data to be parsed.

        Returns:
            dict: Result of parsing.
        """
        if type(data) == str:
            try:
                data = json.loads(data)
                return self.__recursive_parser_str2json(data)
            except Exception:
                pass

        if type(data) == list:
            for i, v in enumerate(data):
                data[i] = self.__recursive_parser_str2json(v)

        if type(data) == dict:
            for k, v in data.items():
                data[k] = self.__recursive_parser_str2json(v)

        return data

    def __recursive_parser_json2str(self, data: Any) -> Any:
        """
        Recursive parser from JSON dict to string.

        Args:
            data (Any): Data to be parsed.

        Returns:
            Any: Data to be written to disk.
        """
        if type(data) == list:
            for i, v in enumerate(data):
                data[i] = self.__recursive_parser_json2str(v)

        if type(data) == dict:
            for k, v in data.items():
                if k in self.special_parsing_keys:
                    data[k] = json.dumps(v)
                else:
                    data[k] = self.__recursive_parser_json2str(v)

        return data

    def extract(self) -> 'PBIXSectionManager':
        """
        Extracts the PBIX file to disk and sets the `layout` instance
        variable.

        Returns:
            PBIXSectionManager: self.
        """
        super().extract()

        layout_path = self.extract_path / "Report" / "Layout"
        with codecs.open(layout_path, encoding=self.encoding) as file:
            str_data = file.read()
            self.layout = self.__recursive_parser_str2json(str_data)

        return self

    def save(self, out_pbix: str) -> None:
        """
        Exports the modified PBIX to disk.

        Args:
            out_pbix (str): Path of the PBIX output file.
        """
        layout_path = self.extract_path / "Report" / "Layout"
        with codecs.open(layout_path, "w", self.encoding) as out_file:
            layout = self.__recursive_parser_json2str(self.layout)
            out_file.write(json.dumps(layout))

        super().save(out_pbix)

    def __filter_by_name(self, section: dict, displayName: str) -> dict:
        """
        Returns True if the section displayName matches the name.
        Otherwise returns False.

        Args:
            section (dict): A section.
            displayName (str): displayName to match.
        """
        return section["displayName"] == displayName

    def rename_section(self, crt_name: str, new_name: str) -> 'PBIXSectionManager':
        """
        Renames a Section.

        Args:
            crt_name (str): Section name to be renamed.
            new_name (str): New name of the Section.

        Returns:
            PBIXSectionManager: self.
        """
        sections = self.layout["sections"]
        sec = next(
            filter(
                lambda x: self.__filter_by_name(x, crt_name),
                sections
            )
        )
        sec["displayName"] = new_name
        return self

    def duplicate_section(self, name_to_dup: str, name_after: str, new_name: str) -> 'PBIXSectionManager':
        """
        Duplicates a Section.

        Args:
            name_to_dup (str): Section name to duplicate.
            name_after (str): Section name after which the duplicated Section will be placed.
            new_name (str): New Name of the duplicated Section.

        Returns:
            PBIXSectionManager: self.
        """
        sections = self.layout["sections"]
        dup = next(
            filter(
                lambda x: self.__filter_by_name(x, name_to_dup),
                sections
            )
        )
        dup = copy.deepcopy(dup)
        dup["displayName"] = new_name

        sec = next(
            filter(
                lambda x: self.__filter_by_name(x, name_after),
                sections
            )
        )
        sec_idx = sec["ordinal"] + 1

        sections.insert(sec_idx, dup)
        for i, s in enumerate(sections[sec_idx:]):
            s["ordinal"] = i + sec_idx

        return self
