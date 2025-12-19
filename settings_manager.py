import json
import os
from datetime import datetime
from typing import Dict, List, Optional

class SettingsManager:
    """Verwaltet gespeicherte Simulationseinstellungen"""
    
    def __init__(self, settings_file: str = "saved_settings.json"):
        self.settings_file = settings_file
        self.settings = self._load_settings()
    
    def _load_settings(self) -> Dict:
        """Lädt gespeicherte Einstellungen aus der JSON-Datei"""
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                return {}
        return {}
    
    def _save_settings(self):
        """Speichert Einstellungen in die JSON-Datei"""
        with open(self.settings_file, 'w', encoding='utf-8') as f:
            json.dump(self.settings, f, indent=2, ensure_ascii=False)
    
    def save_configuration(self, name: str, settings: Dict) -> bool:
        """Speichert eine neue Konfiguration"""
        if not name.strip():
            return False
        
        # Erstelle Eintrag mit Zeitstempel
        self.settings[name] = {
            "settings": settings,
            "created": datetime.now().isoformat(),
            "last_modified": datetime.now().isoformat()
        }
        
        self._save_settings()
        return True
    
    def load_configuration(self, name: str) -> Optional[Dict]:
        """Lädt eine gespeicherte Konfiguration"""
        if name in self.settings:
            return self.settings[name]["settings"]
        return None
    
    def get_all_configurations(self) -> List[str]:
        """Gibt alle gespeicherten Konfigurationsnamen zurück"""
        return list(self.settings.keys())
    
    def rename_configuration(self, old_name: str, new_name: str) -> bool:
        """Benennt eine Konfiguration um"""
        if old_name not in self.settings or not new_name.strip():
            return False
        
        if new_name in self.settings:
            return False  # Name bereits vorhanden
        
        # Verschiebe Einstellungen
        self.settings[new_name] = self.settings[old_name]
        self.settings[new_name]["last_modified"] = datetime.now().isoformat()
        del self.settings[old_name]
        
        self._save_settings()
        return True
    
    def delete_configuration(self, name: str) -> bool:
        """Löscht eine Konfiguration"""
        if name in self.settings:
            del self.settings[name]
            self._save_settings()
            return True
        return False
    
    def update_configuration(self, name: str, settings: Dict) -> bool:
        """Aktualisiert eine bestehende Konfiguration"""
        if name in self.settings:
            self.settings[name]["settings"] = settings
            self.settings[name]["last_modified"] = datetime.now().isoformat()
            self._save_settings()
            return True
        return False
    
    def get_configuration_info(self, name: str) -> Optional[Dict]:
        """Gibt Informationen über eine Konfiguration zurück"""
        if name in self.settings:
            return {
                "name": name,
                "created": self.settings[name]["created"],
                "last_modified": self.settings[name]["last_modified"]
            }
        return None 