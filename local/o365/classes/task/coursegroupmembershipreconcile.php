<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Scheduled task to reconcile Microsoft 365 group membership with Moodle course enrolments.
 *
 * @package local_o365
 * @license http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 * @copyright (C) 2026 onwards Microsoft, Inc. (http://microsoft.com/)
 */

namespace local_o365\task;

use core\task\scheduled_task;
use local_o365\feature\coursesync\main;
use local_o365\feature\coursesync\utils as coursesyncutils;
use local_o365\utils;
use Throwable;

defined('MOODLE_INTERNAL') || die();

require_once($CFG->dirroot . '/local/o365/lib.php');

/**
 * Scheduled task to reconcile Microsoft 365 group owners / members with Moodle course enrolments.
 *
 * Membership is normally pushed to Microsoft 365 by event observers. Those updates are skipped when the user has no Microsoft
 * Entra ID mapping yet, or lost when Microsoft Graph fails. This task is a safety net: for every course that is enabled for sync
 * and connected to a Microsoft 365 group, it computes the full difference between Moodle and Microsoft 365 and applies it.
 *
 * @package     local_o365
 * @subpackage  local_o365\task
 */
class coursegroupmembershipreconcile extends scheduled_task {
    /**
     * Return a human-readable name for this task (shown in the admin task UI).
     *
     * @return string
     */
    public function get_name(): string {
        return get_string('task_coursegroupmembershipreconcile', 'local_o365');
    }

    /**
     * Execute the task.
     *
     * @return bool False if the task exited early because sync is not applicable, true otherwise.
     */
    public function execute(): bool {
        global $DB;

        // Settings checks come first, so that no Microsoft Graph call is made when the task has nothing to do.
        $courseusersyncdirection = get_config('local_o365', 'courseusersyncdirection');
        if ($courseusersyncdirection == COURSE_USER_SYNC_DIRECTION_TEAMS_TO_MOODLE) {
            mtrace('Course user sync direction is Teams to Moodle. Exiting.');
            return false;
        }

        if (coursesyncutils::is_enabled() !== true) {
            mtrace('Course sync is disabled. Exiting.');
            return false;
        }

        if (utils::is_connected() !== true) {
            mtrace('Microsoft 365 is not connected. Exiting.');
            return false;
        }

        $graphclient = coursesyncutils::get_unified_api(__METHOD__);
        if (!$graphclient) {
            mtrace('Could not create a Microsoft Graph API client. Exiting.');
            return false;
        }

        $coursesync = new main($graphclient, true);

        // Match the event-driven sync: SDS course membership is only pushed when "sdssyncenrolmenttosds" is enabled.
        $sdscourseids = [];
        if (!get_config('local_o365', 'sdssyncenrolmenttosds')) {
            $sdscourseids = array_flip($DB->get_fieldset_select('local_o365_objects', 'moodleid', 'type = ?', ['sdssection']));
        }

        $processed = 0;
        $failed = 0;
        $skipped = 0;

        $courses = $this->get_courses_to_reconcile();
        foreach ($courses as $course) {
            if (isset($sdscourseids[$course->id])) {
                $skipped++;
                continue;
            }

            mtrace('Reconciling Microsoft 365 group membership for course #' . $course->id . '...');
            try {
                $result = $coursesync->process_course_team_user_sync_from_moodle_to_microsoft(
                    (int) $course->id,
                    $course->groupobjectid
                );
                if ($result === false) {
                    // Details have already been printed by the sync function.
                    $failed++;
                } else {
                    $processed++;
                }
            } catch (Throwable $e) {
                // Don't let one course stop the reconciliation of the others.
                $failed++;
                mtrace('... Failed to reconcile course #' . $course->id . ': ' . $e->getMessage());
                utils::debug('Exception: ' . $e->getMessage(), __METHOD__, $e);
            }
        }
        $courses->close();

        mtrace(
            'Course group membership reconciliation finished. Courses processed: ' . $processed . ', failed: ' . $failed .
            ', skipped (SDS): ' . $skipped . '.'
        );

        return true;
    }

    /**
     * Get the courses that are enabled for sync and connected to a Microsoft 365 group.
     *
     * @return \moodle_recordset Records with fields id and groupobjectid.
     */
    protected function get_courses_to_reconcile(): \moodle_recordset {
        global $DB;

        $sql = 'SELECT crs.id,
                       obj.objectid AS groupobjectid
                  FROM {course} crs
                  JOIN {local_o365_objects} obj ON obj.type = ? AND obj.subtype = ? AND obj.moodleid = crs.id
                 WHERE crs.id != ?';
        $params = ['group', 'course', SITEID];

        $coursesenabled = coursesyncutils::get_enabled_courses();
        if ($coursesenabled !== true) {
            if (is_array($coursesenabled) && !empty($coursesenabled)) {
                [$coursesinsql, $coursesparams] = $DB->get_in_or_equal($coursesenabled);
                $sql .= ' AND crs.id ' . $coursesinsql;
                $params = array_merge($params, $coursesparams);
            } else {
                $sql .= ' AND 1 = 0';
            }
        }

        $sql .= ' ORDER BY crs.id';

        return $DB->get_recordset_sql($sql, $params);
    }
}
