import { TEXT } from "../constants";

type UnmatchedSurveyModalProps = {
  isOpen: boolean;
  fileNames: string[];
  onClose: () => void;
};

export default function UnmatchedSurveyModal({ isOpen, fileNames, onClose }: UnmatchedSurveyModalProps) {
  if (!isOpen) {
    return null;
  }

  return (
    <div className="modalBackdrop" onClick={onClose}>
      <div className="modalCard" onClick={(event) => event.stopPropagation()}>
        <div className="modalHeader">
          <strong>{TEXT.unmatchedTitle}</strong>
        </div>
        <div className="modalBody">
          {fileNames.length === 0 ? (
            <p className="modalEmpty">{TEXT.unmatchedEmpty}</p>
          ) : (
            <ul className="modalList">
              {fileNames.map((name) => (
                <li key={name}>{name}</li>
              ))}
            </ul>
          )}
        </div>
        <div className="modalFooter">
          <button type="button" onClick={onClose}>
            {TEXT.close}
          </button>
        </div>
      </div>
    </div>
  );
}
