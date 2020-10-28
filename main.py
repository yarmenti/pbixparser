
from src.pbixparser import PBIXSectionManager

if __name__ == "__main__":
    mgr = PBIXSectionManager("dashboard.pbix")

    (
        mgr
        .extract()
        .duplicate_section(
            name_to_dup="Page 1",
            name_after="Page 1",
            new_name="Page Foo"
        )
        .save("modifiedDashboard.pbix")
    )
