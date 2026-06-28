"""Tests for FR-15: mask_text personal-information masking (issue #37).

Covers AC1–AC8.  All tests are pure-Python; no fixtures or external services needed.
"""

from __future__ import annotations

import logging
import re

import pytest

from pptx_md.masking import MASK_TOKEN, MaskingOptions, mask_text

# ---------------------------------------------------------------------------
# AC1: enabled=False -> input returned unchanged
# ---------------------------------------------------------------------------


class TestAc1Disabled:
    """ac1_비활성화_원문반환"""

    def test_ac1_disabled_returns_input_unchanged(self) -> None:
        """enabled=False must return the exact input string without modification."""
        options = MaskingOptions(enabled=False)
        text = "Contact user@example.com or call 010-1234-5678"
        result = mask_text(text, options)
        assert result == text

    def test_ac1_disabled_default_is_false(self) -> None:
        """MaskingOptions default must have enabled=False."""
        options = MaskingOptions()
        assert options.enabled is False

    def test_ac1_disabled_empty_string_unchanged(self) -> None:
        """enabled=False on empty string must return empty string."""
        options = MaskingOptions(enabled=False)
        assert mask_text("", options) == ""


# ---------------------------------------------------------------------------
# AC2: enabled=True + defaults, email -> masked
# ---------------------------------------------------------------------------


class TestAc2EmailMasking:
    """ac2_기본패턴_이메일_마스킹"""

    def test_ac2_email_replaced_with_mask_token(self) -> None:
        """user@example.com must be replaced with MASK_TOKEN when enabled."""
        options = MaskingOptions(enabled=True)
        result = mask_text("Send mail to user@example.com please.", options)
        assert MASK_TOKEN in result
        assert "user@example.com" not in result

    def test_ac2_mask_token_value(self) -> None:
        """MASK_TOKEN must equal '[REDACTED]'."""
        assert MASK_TOKEN == "[REDACTED]"

    def test_ac2_non_pii_text_preserved(self) -> None:
        """Non-PII portions of the string must be preserved after masking."""
        options = MaskingOptions(enabled=True)
        result = mask_text("Hello user@example.com world", options)
        assert "Hello" in result
        assert "world" in result


# ---------------------------------------------------------------------------
# AC3: enabled=True, phone number -> masked
# ---------------------------------------------------------------------------


class TestAc3PhoneMasking:
    """ac3_기본패턴_전화번호_마스킹"""

    def test_ac3_phone_dash_format(self) -> None:
        """010-1234-5678 must be replaced with MASK_TOKEN."""
        options = MaskingOptions(enabled=True)
        result = mask_text("Call 010-1234-5678 for info.", options)
        assert MASK_TOKEN in result
        assert "010-1234-5678" not in result

    def test_ac3_phone_dot_format(self) -> None:
        """010.1234.5678 must also be replaced."""
        options = MaskingOptions(enabled=True)
        result = mask_text("Fax: 02.1234.5678", options)
        assert MASK_TOKEN in result
        assert "02.1234.5678" not in result

    def test_ac3_phone_space_format(self) -> None:
        """010 1234 5678 (space-separated) must be replaced."""
        options = MaskingOptions(enabled=True)
        result = mask_text("Mobile 010 1234 5678 here", options)
        assert MASK_TOKEN in result
        assert "010 1234 5678" not in result


# ---------------------------------------------------------------------------
# AC4: custom pattern + custom text -> masked
# ---------------------------------------------------------------------------


class TestAc4CustomPattern:
    """ac4_커스텀패턴_EMP_코드_마스킹"""

    def test_ac4_custom_pattern_replaces_emp_code(self) -> None:
        """EMP-0042 must be replaced when a custom EMP-\\d{4} pattern is supplied."""
        custom = re.compile(r"EMP-\d{4}")
        options = MaskingOptions(enabled=True, patterns=[custom])
        result = mask_text("Employee EMP-0042 submitted a request.", options)
        assert MASK_TOKEN in result
        assert "EMP-0042" not in result

    def test_ac4_custom_pattern_only_no_default_patterns(self) -> None:
        """When only custom patterns are given, default patterns do not apply."""
        custom = re.compile(r"EMP-\d{4}")
        options = MaskingOptions(enabled=True, patterns=[custom])
        # email should NOT be masked since we replaced the default patterns
        text = "user@example.com EMP-0042"
        result = mask_text(text, options)
        assert "user@example.com" in result
        assert "EMP-0042" not in result

    def test_ac4_multiple_custom_patterns(self) -> None:
        """Multiple custom patterns must all be applied."""
        p1 = re.compile(r"EMP-\d{4}")
        p2 = re.compile(r"SSN-\d{3}")
        options = MaskingOptions(enabled=True, patterns=[p1, p2])
        text = "EMP-0042 and SSN-123 both redacted"
        result = mask_text(text, options)
        assert "EMP-0042" not in result
        assert "SSN-123" not in result
        assert result.count(MASK_TOKEN) == 2


# ---------------------------------------------------------------------------
# AC5: no match -> text unchanged
# ---------------------------------------------------------------------------


class TestAc5NoMatch:
    """ac5_매칭없음_원문반환"""

    def test_ac5_no_pii_returns_original(self) -> None:
        """Text with no PII must be returned unchanged (enabled=True, no match)."""
        options = MaskingOptions(enabled=True)
        text = "No personal information here at all."
        result = mask_text(text, options)
        assert result == text

    def test_ac5_empty_string_no_match(self) -> None:
        """Empty string with enabled=True returns empty string."""
        options = MaskingOptions(enabled=True)
        assert mask_text("", options) == ""

    def test_ac5_whitespace_only_no_match(self) -> None:
        """Whitespace-only string with no PII must be returned unchanged."""
        options = MaskingOptions(enabled=True)
        text = "   \t\n  "
        assert mask_text(text, options) == text


# ---------------------------------------------------------------------------
# AC6 (NFR-06): log must not contain original PII text
# ---------------------------------------------------------------------------


class TestAc6NoPiiInLogs:
    """ac6_NFR06_로그에_원본PII_미출현"""

    def test_ac6_email_not_in_log_records(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Log records emitted during mask_text must not contain the original email."""
        options = MaskingOptions(enabled=True)
        pii = "secret@private.org"
        with caplog.at_level(logging.DEBUG, logger="pptx_md.masking"):
            mask_text(f"Contact {pii} immediately", options)
        for record in caplog.records:
            assert (
                pii not in record.getMessage()
            ), f"PII '{pii}' found in log message: {record.getMessage()!r}"

    def test_ac6_phone_not_in_log_records(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Log records must not contain the original phone number."""
        options = MaskingOptions(enabled=True)
        pii = "010-9876-5432"
        with caplog.at_level(logging.DEBUG, logger="pptx_md.masking"):
            mask_text(f"Call {pii} now", options)
        for record in caplog.records:
            assert (
                pii not in record.getMessage()
            ), f"PII '{pii}' found in log message: {record.getMessage()!r}"

    def test_ac6_log_contains_count_metadata(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        """DEBUG log must mention the match count (metadata-only)."""
        options = MaskingOptions(enabled=True)
        with caplog.at_level(logging.DEBUG, logger="pptx_md.masking"):
            mask_text("user@example.com", options)
        messages = [r.getMessage() for r in caplog.records]
        assert any(
            "1" in m and "match" in m for m in messages
        ), "Expected debug log with match count, got: " + str(messages)


# ---------------------------------------------------------------------------
# AC7: three emails -> all three replaced
# ---------------------------------------------------------------------------


class TestAc7MultipleEmails:
    """ac7_이메일3개_전부치환"""

    def test_ac7_three_emails_all_replaced(self) -> None:
        """Three distinct email addresses in one string must all be replaced."""
        options = MaskingOptions(enabled=True)
        text = "a@foo.com sent to b@bar.org and c@baz.net"
        result = mask_text(text, options)
        assert "a@foo.com" not in result
        assert "b@bar.org" not in result
        assert "c@baz.net" not in result
        assert result.count(MASK_TOKEN) == 3

    def test_ac7_duplicate_email_each_occurrence_replaced(self) -> None:
        """Same email appearing twice must have both occurrences replaced."""
        options = MaskingOptions(enabled=True)
        text = "user@example.com and also user@example.com again"
        result = mask_text(text, options)
        assert "user@example.com" not in result
        assert result.count(MASK_TOKEN) == 2

    def test_ac7_mixed_email_and_phone(self) -> None:
        """Mix of email and phone must each be replaced by MASK_TOKEN."""
        options = MaskingOptions(enabled=True)
        text = "Email user@example.com phone 010-1234-5678"
        result = mask_text(text, options)
        assert "user@example.com" not in result
        assert "010-1234-5678" not in result
        assert result.count(MASK_TOKEN) == 2


# ---------------------------------------------------------------------------
# AC8: same input twice -> same output (determinism)
# ---------------------------------------------------------------------------


class TestAc8Determinism:
    """ac8_동일입력_동일출력_결정성"""

    def test_ac8_same_input_same_output_enabled(self) -> None:
        """Calling mask_text twice with identical args must return identical results."""
        options = MaskingOptions(enabled=True)
        text = "Contact user@example.com or 010-1234-5678"
        result1 = mask_text(text, options)
        result2 = mask_text(text, options)
        assert result1 == result2

    def test_ac8_same_input_same_output_disabled(self) -> None:
        """Determinism also holds when masking is disabled."""
        options = MaskingOptions(enabled=False)
        text = "user@example.com"
        assert mask_text(text, options) == mask_text(text, options)

    def test_ac8_no_side_effects_on_options(self) -> None:
        """mask_text must not mutate the MaskingOptions instance."""
        options = MaskingOptions(enabled=True)
        original_patterns = list(options.patterns)
        mask_text("user@example.com 010-1234-5678", options)
        assert options.patterns == original_patterns
        assert options.enabled is True
