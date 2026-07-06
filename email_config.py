EVENTS = {
    "eligible_for_rewards": {
        "subject": "Congratulations! You Are Eligible | Claim Your Rewards - Google Cloud Study Jams",
        "csv": "event_data/eligible_for_rewards/EligibleParticipantsList.csv",
        "template": "event_data/eligible_for_rewards/email.html",
        "event_dir": "event_data/eligible_for_rewards",
    },
    "access_code_claimed_yes": {
        "subject": "Your Progress Report | Google Cloud Study Jams",
        "csv": "event_data/access_code_claimed_yes/12NovProgressReport.csv",
        "template": "event_data/access_code_claimed_yes/email.html",
        "event_dir": "event_data/access_code_claimed_yes",
    },
    "email_redemption_status_no": {
        "subject": "Action Required: Redeem Your Study Jams Access Code",
        "csv": "event_data/email_redemption_status_no/cloud_study_jams_participants.csv",
        "template": "event_data/email_redemption_status_no/email.html",
        "event_dir": "event_data/email_redemption_status_no",
    },
    "study_jams_goodies_distribution": {
        "subject": "You're Invited! Goodies Distribution Ceremony — Google Cloud Study Jams",
        "csv": "event_data/study_jams_goodies_distribution/test_leaderboard.csv",
        "template": "event_data/study_jams_goodies_distribution/email.html",
        "event_dir": "event_data/study_jams_goodies_distribution",
        "attachments_by_field": {
            "field": "tier",
            "map": {
                "1": ["event_data/study_jams_goodies_distribution/Vouchers/tier-1.png"],
                "2": ["event_data/study_jams_goodies_distribution/Vouchers/tier-2.png"],
                "3": ["event_data/study_jams_goodies_distribution/Vouchers/tier-3.png"],
            },
        },
    },
}
