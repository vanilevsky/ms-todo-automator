import { List, Detail, Toast, showToast, Icon, Form, ActionPanel, Action } from "@raycast/api";
import { useState, useEffect } from "react";
import * as microsoft from "./oauth/microsoft";

const serviceName = "microsoft";

type Values = {
    textfield: string;
    textarea: string;
    datepicker: Date;
    checkbox: boolean;
    dropdown: string;
    tokeneditor: string[];
};

export default function Command() {
    const service = getService(serviceName);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [items, setItems] = useState<{ id: string; title: string }[]>([]);

    useEffect(() => {
        (async () => {
            try {
                await service.authorize();
                // const fetchedItems = await service.fetchItems();
                // setItems(fetchedItems);
                setIsLoading(false);
            } catch (error) {
                console.error(error);
                setIsLoading(false);
                showToast({ style: Toast.Style.Failure, title: String(error) });
            }
        })();
    }, [service]);

    async function handleSubmit(values: Values, ) {
        const createTask = await service.createTask(values);
        console.log(createTask)
    }

    if (!isLoading) {
        return (
            <Form
                isLoading={isLoading}
                actions={
                    <ActionPanel>
                        <Action.SubmitForm onSubmit={handleSubmit} />
                    </ActionPanel>
                }
            >
                <Form.Description text="This form showcases all available form elements." />
                <Form.TextField id="textfield" title="Text field" placeholder="Enter text" defaultValue="Raycast" />
                <Form.TextArea id="textarea" title="Text area" placeholder="Enter multi-line text" />
                <Form.Separator />
                <Form.DatePicker id="datepicker" title="Date picker" />
                <Form.Checkbox id="checkbox" title="Checkbox" label="Checkbox Label" storeValue />
                <Form.Dropdown id="dropdown" title="Dropdown">
                    <Form.Dropdown.Item value="dropdown-item" title="Dropdown Item" />
                </Form.Dropdown>
                <Form.TagPicker id="tokeneditor" title="Tag picker">
                    <Form.TagPicker.Item value="tagpicker-item" title="Tag Picker Item" />
                </Form.TagPicker>
            </Form>
        );
    }

    return (
        <List isLoading={isLoading}>
            {items.map((item) => {
                return <List.Item key={item.id} id={item.id} icon={Icon.TextDocument} title={item.title} />;
            })}
        </List>
    );
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
    fetchItems(): Promise<{ id: string; title: string }[]>;
}
