/**
 * Microsoft Calendar service layer.
 * Provides high-level operations for calendar events.
 */

import { SecureGraphClient } from './client.js';
import { logger } from '../security/logger.js';

/**
 * Microsoft Calendar event attendee.
 */
export interface CalendarAttendee {
  emailAddress: {
    name: string;
    address: string;
  };
  type: 'required' | 'optional' | 'resource';
  status?: {
    response: 'none' | 'organizer' | 'tentativelyAccepted' | 'accepted' | 'declined' | 'notResponded';
  };
}

/**
 * Microsoft Calendar event.
 */
export interface CalendarEvent {
  id: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
  };
  organizer?: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  attendees?: CalendarAttendee[];
  isAllDay: boolean;
  webLink?: string;
}

/**
 * Service for Microsoft Calendar operations.
 */
export class CalendarService {
  constructor(private readonly graphClient: SecureGraphClient) {
    logger.info('CalendarService initialized');
  }

  /**
   * Get calendar events for today.
   *
   * @returns Array of calendar events for today
   */
  async getEventsToday(): Promise<CalendarEvent[]> {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    return this.getCalendarView(today, tomorrow);
  }

  /**
   * Get calendar events for this week (next 7 days).
   *
   * @returns Array of calendar events for the next 7 days
   */
  async getEventsThisWeek(): Promise<CalendarEvent[]> {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 7);

    return this.getCalendarView(today, nextWeek);
  }

  /**
   * Get calendar events for this month (next 30 days).
   *
   * @returns Array of calendar events for the next 30 days
   */
  async getEventsThisMonth(): Promise<CalendarEvent[]> {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const nextMonth = new Date(today);
    nextMonth.setDate(today.getDate() + 30);

    return this.getCalendarView(today, nextMonth);
  }

  /**
   * Get calendar events in a date range using calendarView.
   * The calendarView endpoint automatically expands recurring events.
   *
   * @param startDateTime - Start of the date range
   * @param endDateTime - End of the date range
   * @returns Array of calendar events
   */
  async getCalendarView(startDateTime: Date, endDateTime: Date): Promise<CalendarEvent[]> {
    logger.debug('Fetching calendar view', {
      startDateTime: startDateTime.toISOString(),
      endDateTime: endDateTime.toISOString(),
    });

    const response = await this.graphClient.get<CalendarEvent>('/me/calendar/calendarView', {
      queryParams: {
        startDateTime: startDateTime.toISOString(),
        endDateTime: endDateTime.toISOString(),
        $select: 'id,subject,start,end,location,organizer,attendees,isAllDay,webLink',
        $orderby: 'start/dateTime asc',
        $top: '100',
      },
    });

    const events = response.value ?? [];
    logger.info('Calendar events retrieved', {
      count: events.length,
      startDateTime: startDateTime.toISOString(),
      endDateTime: endDateTime.toISOString(),
    });

    return events;
  }

  /**
   * Get a specific calendar event by ID.
   *
   * @param eventId - Event identifier
   * @returns Calendar event details
   */
  async getEvent(eventId: string): Promise<CalendarEvent> {
    logger.debug('Fetching calendar event', { eventId });

    const response = await this.graphClient.get<CalendarEvent>(
      `/me/events/${eventId}`,
      {
        queryParams: {
          $select: 'id,subject,start,end,location,organizer,attendees,isAllDay,webLink',
        },
      }
    );

    // Single resource response doesn't have 'value' array
    return response as unknown as CalendarEvent;
  }
}
