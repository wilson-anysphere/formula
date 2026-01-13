import type { RibbonTabDefinition } from "../ribbonSchema.js";

export const reviewTab: RibbonTabDefinition = {
  id: "review",
  label: "Review",
  groups: [
    {
      id: "review.proofing",
      label: "Proofing",
      buttons: [
        {
          id: "review.proofing.spelling",
          label: "Spelling",
          ariaLabel: "Spelling",
          iconId: "check",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "review.proofing.spelling", label: "Spelling", ariaLabel: "Spelling", iconId: "check" },
            { id: "review.proofing.spelling.thesaurus", label: "Thesaurus", ariaLabel: "Thesaurus", iconId: "help" },
            { id: "review.proofing.spelling.wordCount", label: "Word Count", ariaLabel: "Word Count", iconId: "hash" },
          ],
        },
        { id: "review.proofing.accessibility", label: "Check Accessibility", ariaLabel: "Check Accessibility", iconId: "help", kind: "dropdown" },
        { id: "review.proofing.smartLookup", label: "Smart Lookup", ariaLabel: "Smart Lookup", iconId: "search", kind: "dropdown" },
      ],
    },
    {
      id: "review.comments",
      label: "Comments",
      buttons: [
        { id: "comments.addComment", label: "New Comment", ariaLabel: "New Comment", iconId: "comment", size: "large" },
        {
          id: "review.comments.deleteComment",
          label: "Delete",
          ariaLabel: "Delete Comment",
          iconId: "trash",
          kind: "dropdown",
          menuItems: [
            { id: "review.comments.deleteComment", label: "Delete Comment", ariaLabel: "Delete Comment", iconId: "trash" },
            { id: "review.comments.deleteComment.deleteThread", label: "Delete Thread", ariaLabel: "Delete Thread", iconId: "comment" },
            { id: "review.comments.deleteComment.deleteAll", label: "Delete All Comments", ariaLabel: "Delete All Comments", iconId: "trash" },
          ],
        },
        { id: "review.comments.previous", label: "Previous", ariaLabel: "Previous Comment", iconId: "arrowUp" },
        { id: "review.comments.next", label: "Next", ariaLabel: "Next Comment", iconId: "arrowDown" },
        { id: "comments.togglePanel", label: "Show Comments", ariaLabel: "Show Comments", iconId: "eye", kind: "toggle" },
      ],
    },
    {
      id: "review.notes",
      label: "Notes",
      buttons: [
        {
          id: "review.notes.newNote",
          label: "New Note",
          ariaLabel: "New Note",
          iconId: "file",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "review.notes.newNote", label: "New Note", ariaLabel: "New Note", iconId: "file" },
            { id: "review.notes.editNote", label: "Edit Note", ariaLabel: "Edit Note", iconId: "edit" },
          ],
        },
        { id: "review.notes.showAllNotes", label: "Show All Notes", ariaLabel: "Show All Notes", iconId: "eye", kind: "toggle" },
        { id: "review.notes.showHideNote", label: "Show/Hide Note", ariaLabel: "Show or Hide Note", iconId: "eyeOff", kind: "toggle" },
      ],
    },
    {
      id: "review.protect",
      label: "Protect",
      buttons: [
        {
          id: "review.protect.protectSheet",
          label: "Protect Sheet",
          ariaLabel: "Protect Sheet",
          iconId: "lock",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "review.protect.protectSheet", label: "Protect Sheet…", ariaLabel: "Protect Sheet", iconId: "lock" },
            { id: "review.protect.unprotectSheet", label: "Unprotect Sheet…", ariaLabel: "Unprotect Sheet", iconId: "unlock" },
          ],
        },
        {
          id: "review.protect.protectWorkbook",
          label: "Protect Workbook",
          ariaLabel: "Protect Workbook",
          iconId: "settings",
          kind: "dropdown",
          menuItems: [
            { id: "review.protect.protectWorkbook", label: "Protect Workbook…", ariaLabel: "Protect Workbook", iconId: "settings" },
            { id: "review.protect.unprotectWorkbook", label: "Unprotect Workbook…", ariaLabel: "Unprotect Workbook", iconId: "unlock" },
          ],
        },
        {
          id: "review.protect.allowEditRanges",
          label: "Allow Edit Ranges",
          ariaLabel: "Allow Edit Ranges",
          iconId: "check",
          kind: "dropdown",
          menuItems: [
            { id: "review.protect.allowEditRanges", label: "Allow Users to Edit Ranges…", ariaLabel: "Allow Users to Edit Ranges", iconId: "check" },
            { id: "review.protect.allowEditRanges.new", label: "New…", ariaLabel: "New allowed range", iconId: "plus" },
          ],
        },
      ],
    },
    {
      id: "review.ink",
      label: "Ink",
      buttons: [
        { id: "review.ink.startInking", label: "Start Inking", ariaLabel: "Start Inking", iconId: "edit", kind: "toggle", size: "large" },
      ],
    },
    {
      id: "review.language",
      label: "Language",
      buttons: [
        {
          id: "review.language.translate",
          label: "Translate",
          ariaLabel: "Translate",
          iconId: "globe",
          kind: "dropdown",
          menuItems: [
            { id: "review.language.translate.translateSelection", label: "Translate Selection", ariaLabel: "Translate Selection", iconId: "globe" },
            { id: "review.language.translate.translateSheet", label: "Translate Sheet", ariaLabel: "Translate Sheet", iconId: "file" },
          ],
        },
        {
          id: "review.language.language",
          label: "Language",
          ariaLabel: "Language",
          iconId: "globe",
          kind: "dropdown",
          menuItems: [
            { id: "review.language.language.setProofing", label: "Set Proofing Language…", ariaLabel: "Set Proofing Language", iconId: "globe" },
            { id: "review.language.language.translate", label: "Translate", ariaLabel: "Translate", iconId: "globe" },
          ],
        },
      ],
    },
    {
      id: "review.changes",
      label: "Changes",
      buttons: [
        {
          id: "review.changes.trackChanges",
          label: "Track Changes",
          ariaLabel: "Track Changes",
          iconId: "edit",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "review.changes.trackChanges", label: "Track Changes…", ariaLabel: "Track Changes", iconId: "edit" },
            { id: "review.changes.trackChanges.highlight", label: "Highlight Changes…", ariaLabel: "Highlight Changes", iconId: "fillColor" },
          ],
        },
        {
          id: "review.changes.shareWorkbook",
          label: "Share Workbook",
          ariaLabel: "Share Workbook",
          iconId: "users",
          kind: "dropdown",
          menuItems: [
            { id: "review.changes.shareWorkbook", label: "Share Workbook…", ariaLabel: "Share Workbook", iconId: "users" },
            { id: "review.changes.shareWorkbook.shareNow", label: "Share Now", ariaLabel: "Share Now", iconId: "link" },
          ],
        },
        {
          id: "review.changes.protectShareWorkbook",
          label: "Protect and Share Workbook",
          ariaLabel: "Protect and Share Workbook",
          iconId: "lock",
          kind: "dropdown",
          menuItems: [
            { id: "review.changes.protectShareWorkbook", label: "Protect and Share Workbook…", ariaLabel: "Protect and Share Workbook", iconId: "lock" },
            { id: "review.changes.protectShareWorkbook.protectWorkbook", label: "Protect Workbook", ariaLabel: "Protect Workbook", iconId: "settings" },
          ],
        },
      ],
    },
  ],
};
