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
 * Tests for Moodle to Microsoft 365 course group membership sync and reconciliation.
 *
 * @package local_o365
 * @license http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 * @copyright (C) 2026 onwards Microsoft, Inc. (http://microsoft.com/)
 */

namespace local_o365;

use advanced_testcase;
use core\context\course;
use core\task\manager;
use local_o365\feature\coursesync\utils;
use local_o365\task\coursegroupmembershipreconcile;

defined('MOODLE_INTERNAL') || die();

global $CFG;

require_once($CFG->dirroot . '/local/o365/lib.php');

/**
 * Tests for Moodle to Microsoft 365 course group membership sync and reconciliation.
 *
 * None of these tests make Microsoft Graph API calls: the plugin is not configured in the test environment, and every code path
 * exercised here returns before a Microsoft Graph API client is created.
 *
 * @group local_o365
 * @covers \local_o365\feature\coursesync\utils
 * @covers \local_o365\task\coursegroupmembershipreconcile
 */
final class coursegroupmembership_test extends advanced_testcase {
    /** @var string Class name of the per-user course group membership sync ad-hoc task. */
    private const ADHOC_TASK = '\local_o365\task\usergroupmembershipsync';

    /**
     * Perform setup before every test. This tells Moodle's phpunit to reset the database after every test.
     */
    protected function setUp(): void {
        parent::setUp();
        $this->resetAfterTest(true);
    }

    /**
     * Create a local_o365_objects record connecting a course to a Microsoft 365 group.
     *
     * @param int $courseid
     * @return int The record ID.
     */
    private function create_course_group_object(int $courseid): int {
        global $DB;

        return $DB->insert_record('local_o365_objects', (object) [
            'type' => 'group',
            'subtype' => 'course',
            'objectid' => 'group-' . $courseid,
            'moodleid' => $courseid,
            'o365name' => 'Course ' . $courseid,
            'timecreated' => time(),
            'timemodified' => time(),
        ]);
    }

    /**
     * Create a local_o365_objects record mapping a Moodle user to a Microsoft Entra ID user.
     *
     * @param int $userid
     * @return int The record ID.
     */
    private function create_user_object(int $userid): int {
        global $DB;

        return $DB->insert_record('local_o365_objects', (object) [
            'type' => 'user',
            'subtype' => '',
            'objectid' => 'user-' . $userid,
            'moodleid' => $userid,
            'o365name' => 'user' . $userid . '@example.onmicrosoft.com',
            'timecreated' => time(),
            'timemodified' => time(),
        ]);
    }

    /**
     * Return the custom data of all queued usergroupmembershipsync ad-hoc tasks, keyed by course ID.
     *
     * @return array
     */
    private function get_queued_sync_tasks(): array {
        $result = [];
        foreach (manager::get_adhoc_tasks(self::ADHOC_TASK) as $task) {
            $data = $task->get_custom_data();
            $result[(int) $data->courseid] = $data;
        }

        return $result;
    }

    /**
     * Test queue_course_group_sync_for_user() only queues tasks for sync-enabled, group-connected courses.
     */
    public function test_queue_course_group_sync_for_user(): void {
        global $DB;

        $generator = $this->getDataGenerator();
        $user = $generator->create_user();
        $connected = $generator->create_course();
        $notconnected = $generator->create_course();
        $syncdisabled = $generator->create_course();
        $notenrolled = $generator->create_course();
        $suspended = $generator->create_course();

        $generator->enrol_user($user->id, $connected->id, 'student');
        $generator->enrol_user($user->id, $notconnected->id, 'student');
        $generator->enrol_user($user->id, $syncdisabled->id, 'student');
        $generator->enrol_user($user->id, $suspended->id, 'student', 'manual', 0, 0, ENROL_USER_SUSPENDED);

        foreach ([$connected, $syncdisabled, $notenrolled, $suspended] as $course) {
            $this->create_course_group_object($course->id);
        }

        // Configure sync after creating test data, so course/enrolment observers have nothing to do.
        set_config('courseusersyncdirection', COURSE_USER_SYNC_DIRECTION_MOODLE_TO_TEAMS, 'local_o365');
        set_config('coursesync', 'oncustom', 'local_o365');
        set_config(
            'coursesynccustom',
            json_encode([$connected->id => 1, $notconnected->id => 1, $notenrolled->id => 1, $suspended->id => 1]),
            'local_o365'
        );
        $DB->delete_records('task_adhoc');

        // Suspended enrolments are included on purpose: the ad-hoc task removes such users from the group.
        $this->assertSame(2, utils::queue_course_group_sync_for_user($user->id));
        $tasks = $this->get_queued_sync_tasks();
        $this->assertEqualsCanonicalizing([$connected->id, $suspended->id], array_keys($tasks));
        $this->assertEquals($user->id, $tasks[$connected->id]->userid);
        $this->assertEquals(0, $tasks[$connected->id]->excluderoleid);

        // Calling again does not create duplicate tasks.
        $this->assertSame(0, utils::queue_course_group_sync_for_user($user->id));
        $this->assertCount(2, $this->get_queued_sync_tasks());

        // With sync on all courses, every enrolled course with a group record is queued.
        $DB->delete_records('task_adhoc');
        set_config('coursesync', 'onall', 'local_o365');
        $this->assertSame(3, utils::queue_course_group_sync_for_user($user->id));
        $this->assertEqualsCanonicalizing(
            [$connected->id, $syncdisabled->id, $suspended->id],
            array_keys($this->get_queued_sync_tasks())
        );

        // Nothing is queued when the sync direction is Teams to Moodle.
        $DB->delete_records('task_adhoc');
        set_config('courseusersyncdirection', COURSE_USER_SYNC_DIRECTION_TEAMS_TO_MOODLE, 'local_o365');
        $this->assertSame(0, utils::queue_course_group_sync_for_user($user->id));
        $this->assertEmpty($this->get_queued_sync_tasks());

        // Both directions behave like Moodle to Teams.
        set_config('courseusersyncdirection', COURSE_USER_SYNC_DIRECTION_BOTH, 'local_o365');
        $this->assertSame(3, utils::queue_course_group_sync_for_user($user->id));

        // Nothing is queued when course sync is off.
        $DB->delete_records('task_adhoc');
        set_config('coursesync', 'off', 'local_o365');
        $this->assertSame(0, utils::queue_course_group_sync_for_user($user->id));
        $this->assertEmpty($this->get_queued_sync_tasks());

        // Nothing is queued for a user without enrolments.
        set_config('coursesync', 'onall', 'local_o365');
        $this->assertSame(0, utils::queue_course_group_sync_for_user($generator->create_user()->id));
    }

    /**
     * Test get_user_group_role_by_moodle_ids() excludes the excluded role correctly when the user has multiple roles.
     */
    public function test_get_user_group_role_by_moodle_ids_with_excluded_role(): void {
        global $DB;

        $generator = $this->getDataGenerator();
        $course = $generator->create_course();
        $context = course::instance($course->id);
        $studentroleid = (int) $DB->get_field('role', 'id', ['shortname' => 'student']);
        $teacherroleid = (int) $DB->get_field('role', 'id', ['shortname' => 'teacher']);
        $editingteacherroleid = (int) $DB->get_field('role', 'id', ['shortname' => 'editingteacher']);

        // User with student (team member), non-editing teacher (no team capability) and editing teacher (team owner) roles.
        $user = $generator->create_user();
        $generator->enrol_user($user->id, $course->id, 'student');
        role_assign($teacherroleid, $user->id, $context->id);
        role_assign($editingteacherroleid, $user->id, $context->id);

        $this->assertSame(MICROSOFT365_GROUP_ROLE_OWNER, utils::get_user_group_role_by_moodle_ids($user->id, $course->id));
        // Excluding the owner role must downgrade to member, regardless of the position of the role in the list.
        $this->assertSame(
            MICROSOFT365_GROUP_ROLE_MEMBER,
            utils::get_user_group_role_by_moodle_ids($user->id, $course->id, $editingteacherroleid)
        );
        $this->assertSame(
            MICROSOFT365_GROUP_ROLE_OWNER,
            utils::get_user_group_role_by_moodle_ids($user->id, $course->id, $studentroleid)
        );
        $this->assertSame(
            MICROSOFT365_GROUP_ROLE_OWNER,
            utils::get_user_group_role_by_moodle_ids($user->id, $course->id, $teacherroleid)
        );

        // User with student and non-editing teacher roles: excluding student leaves no team role.
        $user2 = $generator->create_user();
        $generator->enrol_user($user2->id, $course->id, 'student');
        role_assign($teacherroleid, $user2->id, $context->id);

        $this->assertSame(MICROSOFT365_GROUP_ROLE_MEMBER, utils::get_user_group_role_by_moodle_ids($user2->id, $course->id));
        $this->assertSame('', utils::get_user_group_role_by_moodle_ids($user2->id, $course->id, $studentroleid));
    }

    /**
     * Test sync_user_role_in_course_group_with_status() reports "nothing to do" outcomes without calling Microsoft Graph.
     */
    public function test_sync_user_role_in_course_group_nothing_to_do(): void {
        global $DB;

        $generator = $this->getDataGenerator();
        $user = $generator->create_user();
        $course = $generator->create_course();
        $generator->enrol_user($user->id, $course->id, 'student');

        $this->expectOutputRegex('/no Microsoft Entra ID mapping/');

        $this->assertSame(
            utils::USER_ROLE_SYNC_INVALID_ARGUMENTS,
            utils::sync_user_role_in_course_group_with_status(0, $course->id)
        );

        $this->assertSame(
            utils::USER_ROLE_SYNC_NO_COURSE_GROUP,
            utils::sync_user_role_in_course_group_with_status($user->id, $course->id)
        );
        // Backward compatible wrapper still returns false and does not throw.
        $this->assertFalse(utils::sync_user_role_in_course_group($user->id, $course->id));

        $this->create_course_group_object($course->id);
        $this->assertSame(
            utils::USER_ROLE_SYNC_NO_USER_MAPPING,
            utils::sync_user_role_in_course_group_with_status($user->id, $course->id)
        );

        $this->create_user_object($user->id);
        $DB->insert_record('local_o365_objects', (object) [
            'type' => 'sdssection',
            'subtype' => 'course',
            'objectid' => 'sds-' . $course->id,
            'moodleid' => $course->id,
            'o365name' => 'SDS section',
            'timecreated' => time(),
            'timemodified' => time(),
        ]);
        set_config('sdssyncenrolmenttosds', 0, 'local_o365');
        $this->assertSame(
            utils::USER_ROLE_SYNC_SDS_SYNC_DISABLED,
            utils::sync_user_role_in_course_group_with_status($user->id, $course->id)
        );
    }

    /**
     * Data provider for test_reconcile_task_exits_early.
     *
     * @return array
     */
    public static function reconcile_task_exits_early_provider(): array {
        return [
            'Teams to Moodle direction' => [COURSE_USER_SYNC_DIRECTION_TEAMS_TO_MOODLE, 'onall', '/Teams to Moodle\. Exiting\./'],
            'Course sync off' => [COURSE_USER_SYNC_DIRECTION_MOODLE_TO_TEAMS, 'off', '/Course sync is disabled\. Exiting\./'],
            'Course sync off, both directions' => [COURSE_USER_SYNC_DIRECTION_BOTH, 'off', '/Course sync is disabled\. Exiting\./'],
            'Not connected' => [
                COURSE_USER_SYNC_DIRECTION_MOODLE_TO_TEAMS,
                'onall',
                '/Microsoft 365 is not connected\. Exiting\./',
            ],
        ];
    }

    /**
     * Test the reconcile scheduled task returns early, before any Microsoft Graph client is created, when it has nothing to do.
     *
     * @dataProvider reconcile_task_exits_early_provider
     * @param int $direction Course user sync direction.
     * @param string $coursesync Course sync setting.
     * @param string $expectedoutput Regular expression the task output must match.
     */
    public function test_reconcile_task_exits_early(int $direction, string $coursesync, string $expectedoutput): void {
        $course = $this->getDataGenerator()->create_course();
        $this->create_course_group_object($course->id);

        set_config('courseusersyncdirection', $direction, 'local_o365');
        set_config('coursesync', $coursesync, 'local_o365');

        $this->expectOutputRegex($expectedoutput);
        $task = new coursegroupmembershipreconcile();
        // The early-exit messages are all printed before a Microsoft Graph API client is created.
        $this->assertFalse($task->execute());
    }

    /**
     * Test the reconcile scheduled task only selects courses that are enabled for sync and connected to a group.
     */
    public function test_reconcile_task_course_selection(): void {
        $generator = $this->getDataGenerator();
        $connected = $generator->create_course();
        $notconnected = $generator->create_course();
        $syncdisabled = $generator->create_course();
        $this->create_course_group_object($connected->id);
        $this->create_course_group_object($syncdisabled->id);

        set_config('coursesync', 'oncustom', 'local_o365');
        set_config('coursesynccustom', json_encode([$connected->id => 1, $notconnected->id => 1]), 'local_o365');

        $task = new coursegroupmembershipreconcile();
        $method = new \ReflectionMethod($task, 'get_courses_to_reconcile');

        $courses = $method->invoke($task);
        $selected = [];
        foreach ($courses as $course) {
            $selected[$course->id] = $course->groupobjectid;
        }
        $courses->close();
        $this->assertEquals([$connected->id => 'group-' . $connected->id], $selected);

        set_config('coursesync', 'onall', 'local_o365');
        $courses = $method->invoke($task);
        $selected = [];
        foreach ($courses as $course) {
            $selected[] = $course->id;
        }
        $courses->close();
        $this->assertEqualsCanonicalizing([$connected->id, $syncdisabled->id], $selected);

        set_config('coursesync', 'oncustom', 'local_o365');
        set_config('coursesynccustom', json_encode([]), 'local_o365');
        $courses = $method->invoke($task);
        $this->assertFalse($courses->valid());
        $courses->close();
    }
}
