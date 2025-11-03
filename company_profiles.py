"""
Company profiles manager for the accrual GUI.

Usage:
  - Import CompanyManager in your GUI and call mgr.load()
  - Use mgr.get_profile_names() to populate a company combobox
  - Use mgr.get_current_profile() to read current settings
  - Call mgr.add_profile(...), mgr.edit_profile(...), mgr.delete_profile(name) to manage profiles
Profiles are saved to 'company_profiles.json' in the app working directory.
"""
import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import date, datetime

PROFILES_FILE = "company_profiles.json"

@dataclass
class CompanyProfile:
    name: str
    master: str = ""
    paysheets: str = ""
    admin_map: str = ""
    month_index: int = datetime.now().month - 1
    year: int = datetime.now().year
    pay_dates: List[Tuple[str, float]] = None  # store dates as ISO strings

    def to_dict(self):
        return {
            "name": self.name,
            "master": self.master,
            "paysheets": self.paysheets,
            "admin_map": self.admin_map,
            "month_index": self.month_index,
            "year": self.year,
            "pay_dates": self.pay_dates or [],
        }

    @staticmethod
    def from_dict(d: Dict):
        return CompanyProfile(
            name=d.get("name", ""),
            master=d.get("master", ""),
            paysheets=d.get("paysheets", ""),
            admin_map=d.get("admin_map", ""),
            month_index=int(d.get("month_index", datetime.now().month - 1)),
            year=int(d.get("year", datetime.now().year)),
            pay_dates=d.get("pay_dates", []),
        )

class CompanyManager:
    def __init__(self, path: str = PROFILES_FILE):
        self.path = Path(path)
        self.profiles: Dict[str, CompanyProfile] = {}
        self.current: Optional[str] = None

    def load(self):
        if not self.path.exists():
            self.profiles = {}
            self.current = None
            return
        try:
            with open(self.path, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
            self.profiles = {p["name"]: CompanyProfile.from_dict(p) for p in raw.get("profiles", [])}
            self.current = raw.get("current")
            if self.current not in self.profiles:
                self.current = next(iter(self.profiles), None)
        except Exception:
            self.profiles = {}
            self.current = None

    def save(self):
        try:
            payload = {
                "profiles": [p.to_dict() for p in self.profiles.values()],
                "current": self.current
            }
            with open(self.path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except Exception:
            raise

    def get_profile_names(self) -> List[str]:
        return list(self.profiles.keys())

    def get_profile(self, name: str) -> Optional[CompanyProfile]:
        return self.profiles.get(name)

    def get_current_profile(self) -> Optional[CompanyProfile]:
        return self.get_profile(self.current) if self.current else None

    def set_current(self, name: str):
        if name in self.profiles:
            self.current = name
            self.save()

    def add_profile(self, prof: CompanyProfile):
        if prof.name in self.profiles:
            raise KeyError("Profile already exists")
        self.profiles[prof.name] = prof
        self.current = prof.name
        self.save()

    def edit_profile(self, name: str, prof: CompanyProfile):
        if name not in self.profiles:
            raise KeyError("Profile not found")
        # allow rename
        if prof.name != name:
            # rename key
            del self.profiles[name]
        self.profiles[prof.name] = prof
        self.current = prof.name
        self.save()

    def delete_profile(self, name: str):
        if name in self.profiles:
            del self.profiles[name]
            if self.current == name:
                self.current = next(iter(self.profiles), None)
            self.save()