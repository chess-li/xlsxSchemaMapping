import json
import subprocess

from excel_matcher.models import ColumnProfile, FieldMatchCandidate, StageResult, StageStatus


class CodexExecLLMMatcher:
    def __init__(self, enabled: bool = True, timeout_seconds: int = 60, runner=None):
        self.enabled = enabled
        self.timeout_seconds = timeout_seconds
        self.runner = runner or self._run_codex_exec

    def match(
        self,
        profile: ColumnProfile,
        candidates: list[FieldMatchCandidate],
    ) -> StageResult:
        if not self.enabled:
            return StageResult(
                stage="llm",
                status=StageStatus.SKIPPED,
                reason="llm semantic scoring is disabled",
            )
        if not candidates:
            return StageResult(
                stage="llm",
                status=StageStatus.SKIPPED,
                reason="llm semantic scoring requires candidates",
            )

        prompt = self._build_prompt(profile, candidates)
        try:
            output = self.runner(prompt, self.timeout_seconds)
            data = json.loads(output)
            target_field = data["target_field"]
            semantic_score = _clamp(float(data["semantic_score"]))
        except subprocess.TimeoutExpired as exc:
            return _codex_skipped(f"codex exec timed out: {exc}")
        except Exception as exc:
            return _codex_skipped(f"codex exec failed: {exc}")

        return StageResult(
            stage="llm",
            status=StageStatus.COMPLETED,
            candidates=[
                FieldMatchCandidate(
                    target_field=target_field,
                    score=semantic_score,
                    source="llm",
                    reason=str(data.get("reason", "")),
                )
            ],
        )

    def _build_prompt(
        self,
        profile: ColumnProfile,
        candidates: list[FieldMatchCandidate],
    ) -> str:
        payload = {
            "task": "Select the best standard field for an Excel column.",
            "response_format": {
                "target_field": "candidate target field key",
                "semantic_score": "0.0-1.0",
                "reason": "short explanation",
            },
            "profile": {
                "column_name": profile.column_name,
                "data_type": profile.data_type.value,
                "samples": profile.samples[:5],
            },
            "candidates": [
                {
                    "target_field": candidate.target_field,
                    "score": candidate.score,
                    "source": candidate.source,
                    "reason": candidate.reason,
                }
                for candidate in candidates
            ],
        }
        return json.dumps(payload, ensure_ascii=False, indent=2)

    def _run_codex_exec(self, prompt: str, timeout: int) -> str:
        completed = subprocess.run(
            ["codex", "exec", prompt],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError(
                f"codex exec exited {completed.returncode}: {completed.stderr.strip()}"
            )
        return completed.stdout


def _clamp(score: float) -> float:
    return max(0.0, min(1.0, score))


def _codex_skipped(reason: str) -> StageResult:
    return StageResult(stage="llm", status=StageStatus.SKIPPED, reason=reason)
