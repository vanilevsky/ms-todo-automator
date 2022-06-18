export const msApiBaseUrl = "https://graph.microsoft.com/v1.0";

export interface CreateTaskForm {
    listId: string;
    title: string;
    body: string;
    dueDateTime: string;
    reminderDateTime: string;
}

export type TaskListItem = {
    id: string;
    displayName: string;
    wellknownListName: string;
    isOwner: boolean;
    isShared: boolean;
}