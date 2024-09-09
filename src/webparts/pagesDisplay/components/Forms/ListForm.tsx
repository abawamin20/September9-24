import * as React from "react";
import { useState } from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack,
} from "@fluentui/react";
import PagesService from "../PagesList/PagesService";

export interface IListFormProps {
  pageService: PagesService;
  articleId: string; // Passed from parent component
  title: string; // Passed from parent component
  name: string; // Passed from parent component
  link: string; // Passed from parent component
  hideDialog: () => void;
}

const ListForm: React.FunctionComponent<IListFormProps> = (props) => {
  const [feedbackComments, setFeedbackComments] = useState<string>("");

  const handleSubmit = async () => {
    const formData = {
      Article_x0020_ID: props.articleId,
      Title: props.title,
      Name: props.name,
      LinkColumn: {
        Url: props.link,
        Description: props.name,
      },
      FeedBackComments: feedbackComments,
    };

    try {
      await props.pageService.createListItem(formData, "Feedbacks");
      alert("Feedback created successfully!");
      handleCancel(); // Clear the form and close the dialog
    } catch (error) {
      console.error("Error creating list item: ", error);
      alert("Failed to create item.");
    }
  };

  const handleCancel = () => {
    setFeedbackComments(""); // Clear the feedback field
    props.hideDialog(); // Close the dialog
  };

  return (
    <div>
      <h2>Submit Feedback</h2>

      <TextField
        label="Article Id"
        type="number"
        value={props.articleId}
        readOnly
        style={{
          marginBottom: "10px",
        }}
      />

      <TextField
        label="Title"
        type="text"
        value={props.title}
        readOnly
        style={{
          marginBottom: "10px",
        }}
      />

      <TextField
        label="Link Name"
        type="text"
        value={props.name}
        readOnly
        style={{
          marginBottom: "10px",
        }}
      />
      <TextField
        label="Hypherlink"
        type="text"
        value={props.link}
        readOnly
        style={{
          marginBottom: "10px",
        }}
      />

      <TextField
        label="Feedback Comments"
        multiline
        rows={4}
        value={feedbackComments}
        onChange={(_, value) => setFeedbackComments(value || "")}
        style={{
          marginBottom: "10px",
        }}
      />

      <Stack
        horizontal
        tokens={{ childrenGap: 10 }}
        style={{ marginTop: "10px" }}
      >
        <PrimaryButton text="Submit Feedback" onClick={handleSubmit} />
        <DefaultButton text="Cancel" onClick={handleCancel} />
      </Stack>
    </div>
  );
};

export default ListForm;
