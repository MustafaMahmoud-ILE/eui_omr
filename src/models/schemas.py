from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any

@dataclass
class AnswerKey:
    """Represents the correct answers for a specific exam version/model."""
    version: str
    answers: Dict[int, List[str]]  # Supports multiple correct answers e.g. {1: ["A", "B"], 2: ["C"]}

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'AnswerKey':
        return cls(
            version=data["version"],
            # JSON keys are strings, cast back to int
            answers={int(k): v for k, v in data["answers"].items()}
        )


@dataclass
class GradingResult:
    """Represents the raw extraction of a single OMR sheet page."""
    page_number: int
    student_id: Optional[str]
    version: Optional[str]
    
    # Dictionary mapping question number to the chosen letter(s)
    # E.g. {1: ["C"], 2: [], 3: ["A", "B"]}
    answers: Dict[int, List[str]]
    
    # Validation Flags for Manual Review
    id_error: bool = False
    version_error: bool = False
    question_errors: List[int] = field(default_factory=list) # List of question numbers that need review
    manually_reviewed_questions: List[int] = field(default_factory=list) # Track questions the user manually touched
    is_manual_fix: bool = False
    
    # Cropped OpenCV image numpy arrays (kept in memory, NOT serialised to disk)
    # These are picked up by the UI to show the professor what exactly went wrong.
    _id_crop: Any = field(default=None, repr=False)
    _signature_crop: Any = field(default=None, repr=False)
    _version_crop: Any = field(default=None, repr=False)
    _question_crops: Dict[int, Any] = field(default_factory=dict, repr=False)

    def to_dict(self) -> dict:
        """Serializes standard metadata. Explicitly excludes in-memory image arrays to prevent JSON overhead and errors."""
        return {
            "page_number": int(self.page_number),
            "student_id": self.student_id,
            "version": self.version,
            "answers": self.answers,
            "id_error": bool(self.id_error),
            "version_error": bool(self.version_error),
            "question_errors": [int(q) for q in self.question_errors],
            "manually_reviewed_questions": [int(q) for q in self.manually_reviewed_questions],
            "is_manual_fix": bool(self.is_manual_fix)
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'GradingResult':
        """Reconstructs a result from JSON-compatible dictionary with high fault tolerance."""
        try:
            return cls(
                page_number=int(data.get("page_number", 0)),
                student_id=data.get("student_id"),
                version=data.get("version"),
                # JSON keys are strings, cast back to int
                answers={int(k): v for k, v in data.get("answers", {}).items()},
                id_error=bool(data.get("id_error", False)),
                version_error=bool(data.get("version_error", False)),
                question_errors=[int(q) for q in data.get("question_errors", [])],
                manually_reviewed_questions=[int(q) for q in data.get("manually_reviewed_questions", [])],
                is_manual_fix=bool(data.get("is_manual_fix", False))
            )
        except Exception as e:
            # Re-raise with context for the logger to pick up
            raise ValueError(f"Failed to parse GradingResult: {e}")
