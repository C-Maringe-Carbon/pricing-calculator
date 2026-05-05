#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    Creates the AI_Committee list and seeds all 112 tasks into it.
.PARAMETER SiteUrl
    Full URL of the target SharePoint site.
    e.g. https://carbongroup.sharepoint.com/sites/SystemCommittee
.EXAMPLE
    .\Seed-Lists.ps1 -SiteUrl "https://carbongroup.sharepoint.com/sites/SystemCommittee"
#>
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host "`n=== Carbon Group AI Committee — List Seeder ===" -ForegroundColor Cyan
Write-Host "Target: $SiteUrl`n"

# ── CONNECT ───────────────────────────────────────────────────────────────────
Connect-PnPOnline -Url $SiteUrl -Interactive

# ── CREATE LIST & COLUMNS ─────────────────────────────────────────────────────
$listName = 'AI_Committee'
$existing = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existing) {
    Write-Host "List '$listName' already exists — skipping creation." -ForegroundColor Yellow
} else {
    Write-Host "Creating list '$listName'..." -ForegroundColor Green
    New-PnPList -Title $listName -Template GenericList -OnQuickLaunch | Out-Null

    $columns = @(
        @{ DisplayName='Pillar';         InternalName='Pillar';         Type='Text';     Required=$false },
        @{ DisplayName='PillarID';       InternalName='PillarID';       Type='Text';     Required=$false },
        @{ DisplayName='PillarOrder';    InternalName='PillarOrder';    Type='Number';   Required=$false },
        @{ DisplayName='PillarChampion'; InternalName='PillarChampion'; Type='Text';     Required=$false },
        @{ DisplayName='Project';        InternalName='Project';        Type='Text';     Required=$false },
        @{ DisplayName='Phase';          InternalName='Phase';          Type='Text';     Required=$false },
        @{ DisplayName='ProjectOrder';   InternalName='ProjectOrder';   Type='Number';   Required=$false },
        @{ DisplayName='TaskOrder';      InternalName='TaskOrder';      Type='Number';   Required=$false },
        @{ DisplayName='Done';           InternalName='Done';           Type='Boolean';  Required=$false },
        @{ DisplayName='CompletedDate';  InternalName='CompletedDate';  Type='DateTime'; Required=$false },
        @{ DisplayName='Owner';          InternalName='Owner';          Type='Text';     Required=$false },
        @{ DisplayName='Notes';          InternalName='Notes';          Type='Note';     Required=$false }
    )

    foreach ($col in $columns) {
        Add-PnPField -List $listName `
            -DisplayName $col.DisplayName `
            -InternalName $col.InternalName `
            -Type $col.Type `
            -Required:$col.Required | Out-Null
        Write-Host "  + $($col.DisplayName)" -ForegroundColor DarkGray
    }
    Write-Host "List and columns created.`n" -ForegroundColor Green
}

# ── SEED DATA ─────────────────────────────────────────────────────────────────
Write-Host "Seeding tasks..." -ForegroundColor Cyan

$tasks = @(

    # ═══════════════════════════════════════════════════════════════════════════
    # PILLAR 1 — Dynamics CRM
    # ═══════════════════════════════════════════════════════════════════════════

    # Project 1: Dynamics CRM Build & Configuration  |  Month 1 — Foundation
    @{ Title='Set up Dynamics sandbox environment'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=1 },

    @{ Title='Define and build the client record schema (entities, services, contacts, preferred window, goals)'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=2 },

    @{ Title='Map all entity types: company, trust, individual, SMSF — and configure accordingly'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=3 },

    @{ Title='Connect SharePoint/OneDrive as the document layer'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=4 },

    @{ Title='Define job record fields: job type, status, assigned staff, queries, deadlines, completion state'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=5 },

    @{ Title='Build the onboarding intake form (Content Snare equivalent — native to Dynamics)'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=6 },

    @{ Title='Test data flow: onboarding to client record to job record'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=7 },

    @{ Title="Define what 'job ready to start' means as a system state"
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Dynamics CRM Build & Configuration'; Phase='Month 1 — Foundation'; ProjectOrder=1; TaskOrder=8 },

    # Project 2: Client Data Audit & Migration  |  Month 1 — Foundation
    @{ Title='Audit existing client records across all firms — what exists, where it lives, how clean it is'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Client Data Audit & Migration'; Phase='Month 1 — Foundation'; ProjectOrder=2; TaskOrder=1 },

    @{ Title='Identify which clients have entity structure, preferred contact, and prior year documents accessible'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Client Data Audit & Migration'; Phase='Month 1 — Foundation'; ProjectOrder=2; TaskOrder=2 },

    @{ Title='Define minimum data required per client for the Pre-Start Engine to function'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Client Data Audit & Migration'; Phase='Month 1 — Foundation'; ProjectOrder=2; TaskOrder=3 },

    @{ Title='Build migration plan to move existing client data into Dynamics cleanly'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Client Data Audit & Migration'; Phase='Month 1 — Foundation'; ProjectOrder=2; TaskOrder=4 },

    @{ Title='Identify clients who need to be contacted to fill in missing information'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Client Data Audit & Migration'; Phase='Month 1 — Foundation'; ProjectOrder=2; TaskOrder=5 },

    @{ Title='Plan rollout sequence: which firms go first and in what order'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Client Data Audit & Migration'; Phase='Month 1 — Foundation'; ProjectOrder=2; TaskOrder=6 },

    # Project 3: Healthy, Wealthy & Wise — Dynamics Build  |  Month 3 — Advisory Layer
    @{ Title='Build the Healthy, Wealthy & Wise intake form in Dynamics (digital, shareable with client)'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=1 },

    @{ Title='Configure separate completion for client and spouse/partner'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=2 },

    @{ Title='Build the goal tracking view: business, personal, and family goals — visible to the advisor'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=3 },

    @{ Title='Build the timeline view: 1yr, 3yr, 5yr, 10yr goals side by side'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=4 },

    @{ Title='Configure the shared board between advisor and client: both can see and update before meetings'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=5 },

    @{ Title='Build the action tracking system: actions from each meeting, assigned to client or advisor with due dates'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=6 },

    @{ Title='Make action status visible to both parties (like Planner in Teams)'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=7 },

    @{ Title='Build the meeting history log: what was discussed, decided, and what is outstanding'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=8 },

    @{ Title='Configure advisor prompts: system suggests follow-up points between meetings based on open actions'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=9 },

    @{ Title='Build cross-referral tracking and opportunity flag system'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=10 },

    @{ Title='Build the referral dashboard for partners: who has been referred, to which division, what was the result'
       Pillar='Dynamics CRM'; PillarID='p1'; PillarOrder=1; PillarChampion='TBD'
       Project='Healthy, Wealthy & Wise — Dynamics Build'; Phase='Month 3 — Advisory Layer'; ProjectOrder=3; TaskOrder=11 },

    # ═══════════════════════════════════════════════════════════════════════════
    # PILLAR 2 — Process & System Automation
    # ═══════════════════════════════════════════════════════════════════════════

    # Project 4: Job Scheduling Engine  |  Month 2 — Core Engine
    @{ Title="Build the 'preferred completion window' field into client onboarding"
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=1 },

    @{ Title='Build the job scheduling logic: generate the next job instance based on services and window'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=2 },

    @{ Title='Configure the trigger: 30 days before window opens, automatically initiate the Pre-Start process'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=3 },

    @{ Title='Build the production calendar view: upcoming jobs, status, and assigned team members'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=4 },

    @{ Title='Configure alerts: notify pod leaders when a job has not been initiated within the trigger window'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=5 },

    @{ Title='Test scheduling logic across all job types and entity combinations'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=6 },

    @{ Title='For existing clients: backfill preferred window data and generate first scheduled jobs'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Job Scheduling Engine'; Phase='Month 2 — Core Engine'; ProjectOrder=4; TaskOrder=7 },

    # Project 5: AI-Generated Personalised Client Checklists  |  Month 2 — Core Engine
    @{ Title='Build the checklist generation logic: map job type and entity type to required documents'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=1 },

    @{ Title='Integrate prior year data: pull what was requested and received last year'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=2 },

    @{ Title="Build the 'smart additions' layer: flag known issues from prior year"
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=3 },

    @{ Title='Design the client-facing checklist format — clear, simple, itemised'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=4 },

    @{ Title='Build the automated email/notification with deadline and window commitment language'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=5 },

    @{ Title='Test across: individual, company, trust, and SMSF combinations'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=6 },

    @{ Title='Build the response tracking system: log what has been received and what is outstanding'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='AI-Generated Personalised Client Checklists'; Phase='Month 2 — Core Engine'; ProjectOrder=5; TaskOrder=7 },

    # Project 6: Client Query & Document Collection System  |  Month 2 — Core Engine
    @{ Title='Build the query tracking system in Dynamics: each query is a record with status'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=1 },

    @{ Title='Build the document intake portal: clients upload via a link into SharePoint (not email)'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=2 },

    @{ Title='Configure automatic document tagging: linked to the relevant job and query'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=3 },

    @{ Title='Build the partial response logic: only resend unanswered queries'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=4 },

    @{ Title='Configure automatic reminders if no response after X days'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=5 },

    @{ Title="Build the job status update: all queries resolved — 'ready to start'"
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=6 },

    @{ Title="Build team notification: alert accountant when job moves to 'ready to start'"
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=7 },

    @{ Title='Ensure all responses are visible to team — not locked in one inbox'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=8 },

    @{ Title='Test full loop: send, respond, partial, follow up, complete, job unlocks'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Client Query & Document Collection System'; Phase='Month 2 — Core Engine'; ProjectOrder=6; TaskOrder=9 },

    # Project 7: Pre-Start Validation Check  |  Month 2 — Core Engine
    @{ Title='Build the pre-start validation logic: check all required documents against the checklist'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Pre-Start Validation Check'; Phase='Month 2 — Core Engine'; ProjectOrder=7; TaskOrder=1 },

    @{ Title='Cross-reference received documents against prior year requirements'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Pre-Start Validation Check'; Phase='Month 2 — Core Engine'; ProjectOrder=7; TaskOrder=2 },

    @{ Title='Flag any documents that appear missing or incomplete'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Pre-Start Validation Check'; Phase='Month 2 — Core Engine'; ProjectOrder=7; TaskOrder=3 },

    @{ Title="Generate a 'pre-start summary' for the accountant: what we have, what is confirmed, any flags"
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Pre-Start Validation Check'; Phase='Month 2 — Core Engine'; ProjectOrder=7; TaskOrder=4 },

    @{ Title='Block job from entering production queue until pre-start check is passed'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Pre-Start Validation Check'; Phase='Month 2 — Core Engine'; ProjectOrder=7; TaskOrder=5 },

    @{ Title='Build the override: allow pod leader to manually approve a job if there is a valid reason'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Pre-Start Validation Check'; Phase='Month 2 — Core Engine'; ProjectOrder=7; TaskOrder=6 },

    # Project 8: Document Compilation, Signing & Client Output  |  Month 2-3 — Quality Layer
    @{ Title='Build document compilation logic: bundle all job documents in correct order by entity type'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=1 },

    @{ Title='Build cover letter generator: summarise tax outcomes, refunds/payables, key points for the year'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=2 },

    @{ Title="Build the 'proactive insights' section: flag payroll tax proximity, FBT, Div 7A, and other common issues"
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=3 },

    @{ Title='Integrate signing infrastructure into our own stack (not a third-party cost)'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=4 },

    @{ Title='Configure multi-signature routing: different signatories for different entities in the same group'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=5 },

    @{ Title='Build the signing reminder system: automated reminders until all documents are signed'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=6 },

    @{ Title='Track signing status in Dynamics: visible to the full team'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=7 },

    @{ Title='Notify client manager when signing is complete and lodgement is ready'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Document Compilation, Signing & Client Output'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=8; TaskOrder=8 },

    # Project 9: Quarterly Meeting & Advisory System  |  Month 3 — Advisory Layer
    @{ Title='Build the quarterly meeting scheduler: automatically suggest meeting dates based on client preference'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Quarterly Meeting & Advisory System'; Phase='Month 3 — Advisory Layer'; ProjectOrder=9; TaskOrder=1 },

    @{ Title='Build the advisor prep dashboard: surfaces client goals, open actions, recent job outcomes, and flags'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Quarterly Meeting & Advisory System'; Phase='Month 3 — Advisory Layer'; ProjectOrder=9; TaskOrder=2 },

    @{ Title='Build the shared agenda tool: client and advisor can both add agenda items before the meeting'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Quarterly Meeting & Advisory System'; Phase='Month 3 — Advisory Layer'; ProjectOrder=9; TaskOrder=3 },

    @{ Title='Build the post-meeting summary: advisor logs key points, new actions are created and assigned'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Quarterly Meeting & Advisory System'; Phase='Month 3 — Advisory Layer'; ProjectOrder=9; TaskOrder=4 },

    @{ Title='Configure the follow-up prompt: advisor reminded to follow up on client actions 1-2 times between meetings'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Quarterly Meeting & Advisory System'; Phase='Month 3 — Advisory Layer'; ProjectOrder=9; TaskOrder=5 },

    @{ Title='Train all partners and pod leaders on the advisory model and how to run the meetings'
       Pillar='Process & System Automation'; PillarID='p2'; PillarOrder=2; PillarChampion='TBD'
       Project='Quarterly Meeting & Advisory System'; Phase='Month 3 — Advisory Layer'; ProjectOrder=9; TaskOrder=6 },

    # ═══════════════════════════════════════════════════════════════════════════
    # PILLAR 3 — Integration with Government Agencies
    # ═══════════════════════════════════════════════════════════════════════════

    # Project 10: ATO Lodgement Automation  |  Month 2-3 — Quality Layer
    @{ Title='Map all lodgement types: tax returns (individual, company, trust, SMSF), BAS, payroll, ASIC'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Lodgement Automation'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=10; TaskOrder=1 },

    @{ Title='Define the trigger: once all signatures received, automatically prepare lodgement package'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Lodgement Automation'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=10; TaskOrder=2 },

    @{ Title='Build the lodgement trigger in Dynamics: connects to ATO portal or intermediary'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Lodgement Automation'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=10; TaskOrder=3 },

    @{ Title='Configure lodgement status tracking: pending, submitted, accepted, rejected — visible in Dynamics'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Lodgement Automation'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=10; TaskOrder=4 },

    @{ Title='Build the error handling flow: if ATO rejects, flag to the relevant accountant with reason'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Lodgement Automation'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=10; TaskOrder=5 },

    @{ Title='Build the lodgement confirmation notification: notify client and update job status to complete'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Lodgement Automation'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=10; TaskOrder=6 },

    # Project 11: ATO Compliance Monitoring & Debt Alerts  |  Month 3 — Advisory Layer
    @{ Title='Build integration to pull ATO account balances and debt status per client'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Compliance Monitoring & Debt Alerts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=11; TaskOrder=1 },

    @{ Title='Configure alerts: notify client manager when a client has an outstanding ATO debt or payment plan'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Compliance Monitoring & Debt Alerts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=11; TaskOrder=2 },

    @{ Title='Surface ATO correspondence and notices into the client record in Dynamics'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Compliance Monitoring & Debt Alerts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=11; TaskOrder=3 },

    @{ Title='Build the proactive client notification: alert clients to upcoming payment obligations before they are due'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ATO Compliance Monitoring & Debt Alerts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=11; TaskOrder=4 },

    # Project 12: ASIC & Other Agency Integrations  |  Month 3 — Advisory Layer
    @{ Title='Map all ASIC obligations per client: annual reviews, changes of officeholders, registered addresses'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ASIC & Other Agency Integrations'; Phase='Month 3 — Advisory Layer'; ProjectOrder=12; TaskOrder=1 },

    @{ Title='Build automated ASIC annual review reminder and lodgement workflow'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ASIC & Other Agency Integrations'; Phase='Month 3 — Advisory Layer'; ProjectOrder=12; TaskOrder=2 },

    @{ Title='Identify any other agency integrations required (state revenue offices, payroll tax portals)'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ASIC & Other Agency Integrations'; Phase='Month 3 — Advisory Layer'; ProjectOrder=12; TaskOrder=3 },

    @{ Title='Build the compliance calendar: shows all upcoming lodgement deadlines across all agencies for all clients'
       Pillar='Integration with Government Agencies'; PillarID='p3'; PillarOrder=3; PillarChampion='TBD'
       Project='ASIC & Other Agency Integrations'; Phase='Month 3 — Advisory Layer'; ProjectOrder=12; TaskOrder=4 },

    # ═══════════════════════════════════════════════════════════════════════════
    # PILLAR 4 — Technical Know-How
    # ═══════════════════════════════════════════════════════════════════════════

    # Project 13: Workflow Documentation — All Core Job Types  |  Month 1 — Foundation
    @{ Title='Document the full workflow for: Individual Tax Return'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=1 },

    @{ Title='Document the full workflow for: Company Tax Return + Financials'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=2 },

    @{ Title='Document the full workflow for: Trust Tax Return + Financials'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=3 },

    @{ Title='Document the full workflow for: BAS Lodgement'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=4 },

    @{ Title='Document the full workflow for: SMSF Annual Return'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=5 },

    @{ Title='Document the full workflow for: Payroll processing'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=6 },

    @{ Title='For each job type: list every document required by entity type'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=7 },

    @{ Title="For each job type: define what 'complete information' looks like before production starts"
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=8 },

    @{ Title='For each job type: document the full review checklist (what reviewers currently check)'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=9 },

    @{ Title='Identify variations: what changes based on client complexity or special circumstances'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Workflow Documentation — All Core Job Types'; Phase='Month 1 — Foundation'; ProjectOrder=13; TaskOrder=10 },

    # Project 14: AI First-Level Review System  |  Month 2-3 — Quality Layer
    @{ Title='Define the first-level AI review checklist for each job type'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=1 },

    @{ Title='For tax returns: check figures reconcile, prior year comparisons within range, all entities included'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=2 },

    @{ Title='For financials: check balance sheet balances, P&L is complete, comparatives are correct'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=3 },

    @{ Title='For BAS: check GST figures reconcile to coding, period is correct, lodgement date noted'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=4 },

    @{ Title="Build the AI review layer: runs automatically when accountant marks job as 'complete'"
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=5 },

    @{ Title='Generate an AI review report: list of checks passed, any flags for human attention'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=6 },

    @{ Title='Build the workflow gate: job cannot go to human review until AI review is complete'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=7 },

    @{ Title='For tax compliance jobs: compare final outcome to tax planning engagement — flag if more than 5% variance'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=8 },

    @{ Title='Test AI review across all job types and validate accuracy with a senior accountant'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI First-Level Review System'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=14; TaskOrder=9 },

    # Project 15: Human Review Workflow (Technical + Commercial)  |  Month 2-3 — Quality Layer
    @{ Title='Define the technical review checklist: accuracy, completeness, compliance'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=1 },

    @{ Title='Define the commercial review checklist: best tax position, missed opportunities, client flags'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=2 },

    @{ Title='Build the two-stage review workflow in Dynamics: technical to commercial to approved'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=3 },

    @{ Title='Configure routing: simple jobs go to client manager, complex jobs escalate to pod leader or partner'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=4 },

    @{ Title='Build the feedback loop: reviewer sends job back with specific comments, not just a general rejection'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=5 },

    @{ Title='Track review turnaround time: flag jobs sitting in review queue for more than X days'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=6 },

    @{ Title='Build the approval gate: job cannot be compiled for client until both reviews are complete'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='Human Review Workflow (Technical + Commercial)'; Phase='Month 2-3 — Quality Layer'; ProjectOrder=15; TaskOrder=7 },

    # Project 16: AI-Assisted Meeting Preparation & Cross-Referral Prompts  |  Month 3 — Advisory Layer
    @{ Title='Configure AI-generated meeting prompts: based on Healthy, Wealthy & Wise data, suggest relevant topics'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI-Assisted Meeting Preparation & Cross-Referral Prompts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=16; TaskOrder=1 },

    @{ Title='Build the cross-referral prompt layer: surface opportunities to refer into other divisions based on client goals'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI-Assisted Meeting Preparation & Cross-Referral Prompts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=16; TaskOrder=2 },

    @{ Title='Build the gap analysis: which clients have only one service? Flag as expansion opportunities'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI-Assisted Meeting Preparation & Cross-Referral Prompts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=16; TaskOrder=3 },

    @{ Title='Build the revenue attribution model: track cross-referral revenue back to the originating advisor'
       Pillar='Technical Know-How'; PillarID='p4'; PillarOrder=4; PillarChampion='TBD'
       Project='AI-Assisted Meeting Preparation & Cross-Referral Prompts'; Phase='Month 3 — Advisory Layer'; ProjectOrder=16; TaskOrder=4 }
)

$total = $tasks.Count
$i = 0
foreach ($task in $tasks) {
    $i++
    $preview = if ($task.Title.Length -gt 60) { $task.Title.Substring(0, 60) + '...' } else { $task.Title }
    Write-Progress -Activity "Seeding AI_Committee" `
        -Status "$i / $total  |  $preview" `
        -PercentComplete ([int]($i / $total * 100))

    Add-PnPListItem -List $listName -Values @{
        Title          = $task.Title
        Pillar         = $task.Pillar
        PillarID       = $task.PillarID
        PillarOrder    = $task.PillarOrder
        PillarChampion = $task.PillarChampion
        Project        = $task.Project
        Phase          = $task.Phase
        ProjectOrder   = $task.ProjectOrder
        TaskOrder      = $task.TaskOrder
        Done           = $false
    } | Out-Null
}

Write-Progress -Activity "Seeding AI_Committee" -Completed
Write-Host "`n  $total tasks seeded into '$listName'." -ForegroundColor Green
Write-Host "  Next: upload carbon_group_committee.html to Site Assets and add the web part to a page.`n"
