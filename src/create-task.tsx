import { Toast, showToast, Form, ActionPanel, Action, Detail, showHUD } from "@raycast/api";
import { useState, useEffect } from "react";
import * as microsoft from "./oauth/microsoft";
import { CreateTaskForm, TaskListItem } from "./const";

const serviceName = "microsoft";

const defaultListItem = {
    id: "default-list-item-id",
    displayName: "👀",
    wellknownListName: "defaultList",
} as TaskListItem;

export default function Command() {
    const service = getService(serviceName);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [lists, setLists] = useState<TaskListItem[]>([defaultListItem]);

    useEffect(() => {
        (async () => {
            try {
                await service.authorize();

                service.fetchLists()
                    .then((lists) => {
                        setLists(lists.sort((l) => l.wellknownListName === "defaultList" ? -1 : 1));
                    });

                setIsLoading(false);
            } catch (error) {
                console.error(error);
                setIsLoading(false);
                showToast({ style: Toast.Style.Failure, title: String(error) });
            }
        })();
    }, [service]);

    async function handleSubmit(values: CreateTaskForm) {

        if (!values.title) {
            await showToast({
                style: Toast.Style.Failure,
                title: "Task is required",
            });
            return;
        }

        const taskStatus = service.createTask(values);
        await showToast({ style: Toast.Style.Animated, title: "Task is in progress..." });
        taskStatus.then(() => { showHUD("👏 Task created") });
    }

    if (!isLoading) {
        return (
            <Form
                actions={
                    <ActionPanel>
                        <Action.SubmitForm
                            title="Create Task"
                            onSubmit={handleSubmit}
                        />
                    </ActionPanel>
                }
            >
                <Form.TextField id="title" title="Task" placeholder="Add a Task" autoFocus />
                <Form.DatePicker id="dueDateTime" title="Due Date" />
                <Form.DatePicker id="reminderDateTime" title="Reminder" />
                <Form.TextArea id="body" title="Note" placeholder="Add Note" />
                <Form.Dropdown id="listId" title="List">
                    {lists.map((list) => (
                        <Form.Dropdown.Item
                            key={list.id}
                            value={list.id}
                            title={list.displayName}
                        />
                    ))}
                </Form.Dropdown>
            </Form>
        );
    }

    return <Detail />;
}

// Services

function getService(serviceName: string): Service {
    switch (serviceName) {
        case "microsoft":
            return microsoft as Service;
        default:
            throw new Error("Unsupported service: " + serviceName);
    }
}

interface Service {
    authorize(): Promise<void>;
    fetchLists(): Promise<TaskListItem[]>;
    createTask(values: CreateTaskForm): Promise<void>;
}
