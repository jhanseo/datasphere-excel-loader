# -*- coding: utf-8 -*-
"""DRM(DRMONE) 파일 감지와 실행 차단."""

import pytest

import excel_loader as core


def test_drm_file_is_detected(drm_xlsx, sample_xlsx):
    assert core.is_drmone_file(drm_xlsx) is True
    assert core.is_drmone_file(sample_xlsx) is False


def test_missing_file_is_not_reported_as_drm(tmp_path):
    assert core.is_drmone_file(str(tmp_path / "nope.xlsx")) is False


def test_execute_path_refuses_drm_file(drm_xlsx, options, logger):
    """allow_drm=False('그냥 실행')에서는 DrmProtectedError를 던진다."""
    with pytest.raises(core.DrmProtectedError) as excinfo:
        core.collect_sheet_plans([drm_xlsx], options, logger, allow_drm=False)
    assert excinfo.value.file_name == "drm_secret.xlsx"
    assert "미리보기" in str(excinfo.value)


def test_preview_path_does_not_raise_drm_error(drm_xlsx, options, logger):
    """미리보기 경로는 DRM을 거부하지 않는다.

    Windows + Excel이 없는 환경에서는 파일 단위 오류로 수집되고,
    DrmProtectedError는 발생하지 않아야 한다.
    """
    plans, errors = core.collect_sheet_plans([drm_xlsx], options, logger,
                                            allow_drm=True)
    assert plans == []
    assert errors and errors[0][0] == "drm_secret.xlsx"


def test_guide_message_names_the_file():
    message = core.DRM_GUIDE_MESSAGE.format(name="비밀.xlsx")
    assert "비밀.xlsx" in message
    assert "미리보기" in message


def test_non_excel_content_raises_clear_error(tmp_path, options, logger):
    path = tmp_path / "fake.xlsx"
    path.write_text("id,name\n1,a\n", encoding="utf-8")
    plans, errors = core.collect_sheet_plans([str(path)], options, logger,
                                            allow_drm=False)
    assert plans == []
    assert "Excel 파일 내부 형식이 아닙니다" in errors[0][2]
