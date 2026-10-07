"""Timestamped side copy of a deck before a risky edit. No COM, so it never disturbs an open
PowerPoint session; save in PowerPoint first, because this copies what is on disk.

    python backup_snapshot.py --file deck.pptx --label pre-bulk-fontbump
    -> deck.20261007-143000.pre-bulk-fontbump.pptx (beside the deck)
"""
import argparse
import datetime
import pathlib
import shutil

from kShared import ToolInputException, kRun, kS, kToolException


class kBackupSnapshot:
    """One snapshot of one deck file."""

    def __init__(self, Source, Label):
        try:
            self.Source = pathlib.Path(Source).resolve()
            self.Label = Label
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBackupSnapshot.__init__")

    def Target(self):
        """The snapshot path beside the deck."""
        if kS.ErrorMode:
            return None
        try:
            Stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            return self.Source.with_name(f"{self.Source.stem}.{Stamp}.{self.Label}{self.Source.suffix}")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBackupSnapshot.Target")
            return None

    def Save(self):
        """Copy the deck; never overwrite an earlier snapshot. Returns the new path."""
        if kS.ErrorMode:
            return None
        try:
            if not self.Source.is_file():
                raise ToolInputException(f"not found: {self.Source}")
            Destination = self.Target()
            if Destination.exists():
                raise ToolInputException(f"refusing to overwrite {Destination}")
            shutil.copy2(self.Source, Destination)
            return Destination
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kBackupSnapshot.Save(file={self.Source})")
            return None


class kBackupSnapshotApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser()
            Parser.add_argument("--file", required=True)
            Parser.add_argument("--label", default="backup")
            Args = Parser.parse_args()
            Saved = kBackupSnapshot(Args.file, Args.label).Save()
            if Saved:
                print(Saved)
            return 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBackupSnapshotApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kBackupSnapshotApp)
