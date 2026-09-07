from conversion_engines import available_engines, find_libreoffice, find_microsoft_office


def test_engine_detection_returns_valid_values() -> None:
    engines = available_engines()
    assert set(engines).issubset({"office", "libreoffice"})
    assert ("office" in engines) == find_microsoft_office()
    assert ("libreoffice" in engines) == (find_libreoffice() is not None)
