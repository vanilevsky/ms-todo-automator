import { Toast, showToast, Form, ActionPanel, Action, Detail } from "@raycast/api";
import { useState, useEffect } from "react";
import * as microsoft from "./oauth/microsoft";
import { CreateTaskForm, TaskListItem } from "./const";

const serviceName = "microsoft";

export default function Command() {
    const service = getService(serviceName);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [lists, setLists] = useState<TaskListItem[]>([]);

    useEffect(() => {
        (async () => {
            try {
                await service.authorize();

                const fetchedLists = await service.fetchLists();
                setLists(fetchedLists);

                console.debug(fetchedLists);

                setIsLoading(false);
            } catch (error) {
                console.error(error);
                setIsLoading(false);
                showToast({ style: Toast.Style.Failure, title: String(error) });
            }
        })();
    }, [service]);

    async function handleSubmit(values: CreateTaskForm) {

        await service.createTask(values);
        await showToast({ style: Toast.Style.Success, title: "Task created" });

    }

    function getDefaultId(lists: TaskListItem[]): string | undefined {
        return lists.find(list => list.wellknownListName === "defaultList")?.id
    }

    if (!isLoading) {
        return (
            <Form
                actions={
                    <ActionPanel>
                        <Action.SubmitForm
                            onSubmit={handleSubmit}
                        />
                    </ActionPanel>
                }
            >
                <Form.TextField id="title" title="Full Name" placeholder="Add a Task" info="Hello" autoFocus />
                <Form.TextArea id="body" title="Full Name" placeholder="Add Note" />
                <Form.Dropdown id="listId" title="Full Name" defaultValue={getDefaultId(lists)}>
                    {lists.map((list) => (
                        <Form.Dropdown.Item
                            key={list.id}
                            value={list.id}
                            title={list.displayName}
                        />
                    ))}
                </Form.Dropdown>
                <Form.DatePicker id="dueDateTime" title="Add Due Date" />
                <Form.DatePicker id="reminderDateTime" title="Add Reminder Date & Time" />
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
