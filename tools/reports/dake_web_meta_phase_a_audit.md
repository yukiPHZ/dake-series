# DAKE_WEB_META Phase A Audit

Generated: 2026-06-02 07:32:05

## Purpose

Read-only audit for DAKE_WEB_META Phase A. This report checks what exists now and does not modify any site README. No site repository commit or push was performed.

## Phase A Report Existence

| file | exists | last_modified | rows | summary |
| --- | --- | --- | --- | --- |
| dake_web_sites_review.md | true | 2026-06-02 06:42:19 |  | site_body_count=22; candidate_count=22; git_not_available_count=0; git_dirty_repo_count=0; url_known_or_readme_count=17; url_inferred_count=4; url_unknown_count=1; missing_DAKE_WEB_META_count=22; existing_DAKE_WEB_META_count=0; openai_reference_site_count=9; health_api_site_count=9; borinef_expected_url=https://borinef.com/ |
| dake_web_sites_review.csv | true | 2026-06-02 06:42:18 | 22 |  |
| dake_web_meta_phase_a_targets.md | true | 2026-06-02 06:48:25 |  | phase_a_target_count=17; known_count=16; readme_count=1; functions_count=7; openai_reference_count=7 |
| dake_web_meta_phase_a_targets.csv | true | 2026-06-02 06:48:25 | 17 |  |
| dake_web_meta_phase_a_result.md | true | 2026-06-02 06:53:21 |  | phase_a_target_count=17; readme_added_count=17; safety_ok_count=17; safety_review_count=0; site_repo_push_done=false; borinef_production_url=https://borinef.com/ |
| dake_web_meta_phase_a_result.csv | true | 2026-06-02 06:53:21 | 17 |  |

## Summary

| metric | value |
| --- | --- |
| audit_record_count | 34 |
| phase_a_target_count | 17 |
| phase_a_result_added_count | 17 |
| all_records_has_DAKE_WEB_META | 17 |
| all_records_missing_DAKE_WEB_META | 17 |
| phase_a_targets_has_DAKE_WEB_META | 17 |
| phase_a_targets_missing_DAKE_WEB_META | 0 |
| done | 1 |
| changed_uncommitted | 0 |
| committed_not_pushed | 16 |
| not_done | 0 |
| human_review | 17 |

## Dashboard Launch Check

```text
LAUNCH CHECK OK: sites=26 candidates=8 component_sites=0 active=17 url_missing=9 needs_organizing=15 api_sites=11 dirty_sites=0 git_errors=6
```

## Interpretation

- Phase A targets/result reports exist.
- Phase A target count is 17; result CSV records 17 rows with `added=True`.
- Current README files also show `DAKE_WEB_META` in 17 of 17 Phase A targets.
- Phase A targets are not uncommitted README edits. Most target repos are clean but `ahead 1`, so the main state is `committed_not_pushed`.
- Dashboard `url_missing` / `needs_organizing` / `git_errors` include more than Phase A targets: inferred URLs, unknown URLs, and non-repo candidates are part of the current dashboard scan.
- Rows without `DAKE_WEB_META` are mainly Phase A excluded inferred/unknown/non-site candidates. They need human review before any README update.

## BORINEF Check

| folder_name | domain | production_url | url_status | phase_a_status | git_status | reason |
| --- | --- | --- | --- | --- | --- | --- |
| borinef | borinef.com | https://borinef.com/ | ok | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True |

BORINEF is read as `domain=borinef.com` and `production_url=https://borinef.com/`. The README contains note.com links, but DAKE_WEB_META `production_url` is not note.com.

## Audit Rows

| folder_name | has_meta | domain | production_url | source | phase_a_status | git | reason | next |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| booth_ready | false |  |  | unknown | human_review | not_git_repo | domain / production_url missing; URL source is unknown; not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| borinef | true | borinef.com | https://borinef.com/ | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| dake-ai-site | true | ai.dakeapp.com | https://ai.dakeapp.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| dake-gis-site | true | gis.dakeapp.com | https://gis.dakeapp.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| dake-labs-site | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| dake-tools-site | true | tools.dakeapp.com | https://tools.dakeapp.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| dakeapp-site | true | dakeapp.com | https://dakeapp.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| holiday-blue-site | true | blue.holiday-jinja.com | https://blue.holiday-jinja.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| holiday-jinja-site | true | holiday-jinja.com | https://holiday-jinja.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| holiday-side-site | true | side.holiday-jinja.com | https://side.holiday-jinja.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| holiday-sky-site | true | sky.holiday-jinja.com | https://sky.holiday-jinja.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| invoice | false | tools.dakeapp.com | https://tools.dakeapp.com/invoice/` | readme | human_review | not_git_repo | not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| japanmemorylane-site | true | japanmemorylane.com | https://japanmemorylane.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| kadochu | false | peakheadz.booth.pm | https://peakheadz.booth.pm/ | readme | human_review | not_git_repo | not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| logs | false |  |  | unknown | human_review | not_git_repo | domain / production_url missing; URL source is unknown; not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| nicekip-restore | true | nicekip-restore.pages.dev | https://nicekip-restore.pages.dev | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| nicekip-site | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| niceskill-site | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| pdf_to_jpeg_app | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| pdf_to_jpeg_app | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| peakheadz-project-index | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| peakheadz-site | true | peakheadz.com | https://peakheadz.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| SHIMARISU | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| shimarisu-dakeapp-site | true | shimarisu.dakeapp.com | https://shimarisu.dakeapp.com | meta | done | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | No Phase A action needed |
| shimarisu-site | true | shimarisu-fudosan.com | https://shimarisu-fudosan.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| soredake-site | true | soredake.com | https://soredake.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| tsucho | false |  |  | unknown | human_review | not_git_repo | domain / production_url missing; URL source is unknown; not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| wlzphz-site | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| yukihikokikuta-site | true | yukihikokikuta.com | https://yukihikokikuta.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| yukizblog-restore | false |  |  | unknown | human_review | clean | domain / production_url missing; URL source is unknown | Confirm URL/repo before adding DAKE_WEB_META |
| yukizblog-site | true | blog.yukihikokikuta.com | https://blog.yukihikokikuta.com | meta | committed_not_pushed | clean | Phase A target has DAKE_WEB_META; Phase A result recorded as added=True | Push the existing site repo commit if ready |
| （凍結）DAKE_PDF_Split | false |  |  | unknown | human_review | not_git_repo | domain / production_url missing; URL source is unknown; not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| （凍結）DAKE_Web_EntryBuilder | false |  |  | unknown | human_review | not_git_repo | domain / production_url missing; URL source is unknown; not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |
| （凍結）shimarisu-fudosan | false |  |  | unknown | human_review | not_git_repo | domain / production_url missing; URL source is unknown; not_git_repo | Confirm URL/repo before adding DAKE_WEB_META |

## Recommended Next Split

- `committed_not_pushed`: README update and local commit likely already happened. If intentional, the next step is pushing the affected site repos.
- `human_review`: Inferred URL, unknown URL, non-repo candidate, or git unavailable. Confirm manually before adding DAKE_WEB_META.
- `not_done`: Known URL with missing DAKE_WEB_META. This audit found no such rows.
