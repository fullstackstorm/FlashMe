AFS_NCI = {
    "Operator Follow-up Miss" : [
        "AA didn’t send an accurate script or no script at all prior to closing the case when the vehicle was not Amazon owned or Amazon leased",
        "AA didn’t send an accurate script or didn’t send one at all when there was an active DVIC for the vehicle after the date/time case was opened",
        "AA didn’t send the correct script when the vehicle location field in salesforce was completed"
    ],
    "False Resolution Miss" : [
        "AA didn’t update the case status accurately when there was an active DVIC for the vehicle after the date/time case was opened",
        "AA didn’t update the status accurately when the vehicle was not at the vendor with the appointment already scheduled",
        "AA didn’t accurately update the scorecard week number",
        "AA didn’t accurately update the subject week in salesforce",
        "AA didn’t accurately update the Case Resolution Barrier field and Case delay cause fields or didn’t update them at all when the case description mentioned parts on back order or dealership capacity issues"
    ]
}

Ungrounding_NCI = {
    "False Resolution Miss" : [
        "AA rejected the automatic approval requests regardless of the document submitted by the DSP",
        "AA Approved with the date of repair earlier than 14 days (exceptions will exist for greater than 14 days out. It should be notated in the notes section)",
        "AA Approved with a certification form with both certifications selected",
        "AA Approved with a certification form with the first option selected",
        "AA Approved with missing DSP Info",
        "AA Approved with no checkbox next to the defect selected in error/one was written in",
        "AA Approved with no proof of repair completion for the AMERIT/Repairsmith/vendor invoices documentations",
        "AA Approved with a Rivian RO when type of repair doesn’t match",
        "AA Approved DOT marked in error when the VIN appeared on the monthly audit List"
    ]
}
