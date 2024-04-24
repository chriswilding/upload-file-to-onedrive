package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"

	"github.com/Azure/azure-sdk-for-go/sdk/azcore"
	"github.com/Azure/azure-sdk-for-go/sdk/azcore/policy"
	msgraphsdk "github.com/microsoftgraph/msgraph-sdk-go"
	"github.com/microsoftgraph/msgraph-sdk-go-core/authentication"
	"github.com/microsoftgraph/msgraph-sdk-go/models"
	"github.com/microsoftgraph/msgraph-sdk-go/models/odataerrors"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/microsoft"
)

type secrets struct {
	ClientID     string `json:"client_id"`
	ClientSecret string `json:"client_secret"`
}

func getSecrets() (secrets, error) {
	secretsFile, err := os.Open("./secrets.json")
	if err != nil {
		return secrets{}, err
	}
	defer secretsFile.Close()

	var s secrets

	err = json.NewDecoder(secretsFile).Decode(&s)
	if err != nil {
		return secrets{}, err
	}

	return s, nil
}

type azureTokenCredential struct {
	token oauth2.TokenSource
}

func (c azureTokenCredential) GetToken(_ context.Context, _ policy.TokenRequestOptions) (azcore.AccessToken, error) {
	t, err := c.token.Token()
	if err != nil {
		return azcore.AccessToken{}, err
	}
	return azcore.AccessToken{
		Token:     t.AccessToken,
		ExpiresOn: t.Expiry,
	}, nil
}

func main() {
	secrets, err := getSecrets()
	if err != nil {
		fmt.Printf("Error getting secrets: %v\n", err)
		return
	}

	config := &oauth2.Config{
		ClientID:     secrets.ClientID,
		ClientSecret: secrets.ClientSecret,
		Endpoint:     microsoft.AzureADEndpoint(""),
		RedirectURL:  "http://localhost:56426",
		Scopes:       []string{"offline_access", "Files.ReadWrite.All"},
	}

	var oauth2Token *oauth2.Token
	f, err := os.Open("./token.json")
	if errors.Is(err, os.ErrNotExist) {
		url := config.AuthCodeURL("1234567890", oauth2.AccessTypeOffline, oauth2.ApprovalForce)
		fmt.Println(url)

		var code string
		fmt.Scan(&code)

		oauth2Token, err = config.Exchange(context.Background(), code)
		if err != nil {
			fmt.Printf("Error exchanging oauth token: %v\n", err)
			return
		}

		jsonData, err := json.Marshal(oauth2Token)
		if err != nil {
			fmt.Printf("Error marshalling oauth token: %v\n", err)
			return
		}

		err = os.WriteFile("token.json", jsonData, 0644)
		if err != nil {
			fmt.Printf("Error writing oauth token: %v\n", err)
			return
		}

		f, err = os.Open("./token.json")
		if err != nil {
			fmt.Printf("Error opening token file: %v\n", err)
			return
		}
	} else if err != nil {
		fmt.Printf("Error opening token file: %v\n", err)
		return
	}

	err = json.NewDecoder(f).Decode(&oauth2Token)
	if err != nil {
		fmt.Printf("Error decoding token file: %v\n", err)
		return
	}
	f.Close()

	atc := azureTokenCredential{
		token: config.TokenSource(context.Background(), oauth2Token),
	}

	tokenProvider, err := authentication.NewAzureIdentityAuthenticationProviderWithScopes(atc, []string{"Files.Read"})
	if err != nil {
		fmt.Printf("Error creating NewAzureIdentityAccessTokenProvider: %v\n", err)
		return
	}

	adapter, err := msgraphsdk.NewGraphRequestAdapter(tokenProvider)
	if err != nil {
		fmt.Printf("Error with NewGraphRequestAdapter: %v\n", err)
		return
	}

	client := msgraphsdk.NewGraphServiceClient(adapter)

	// Get the Drive ID of a Microsoft 365 user, or use "me" for personal.
	// usersDrive, err := client.Me().Drive().Get(context.Background(), nil)
	// if err != nil {
	// 	fmt.Printf("Error getting users drive ID: %v\n", err)
	// 	return
	// }
	// driveID := *usersDrive.GetId()
	driveID := "me"

	createdDriveItem, err := createFolder(client, driveID, "root", "ExampleFolder")
	if err != nil {
		fmt.Printf("Error creating folder item: %v\n", err)
		printOdataError(err)
		return
	}

	createdDriveItem, err = createFolder(client, driveID, *createdDriveItem.GetId(), "ExampleChildFolder")
	if err != nil {
		fmt.Printf("Error creating child folder item: %v\n", err)
		printOdataError(err)
		return
	}

	name := "example.txt"
	contentType := "text/csv"
	data := []byte("Hello,World!")

	newFile := models.NewFile()
	newFile.SetMimeType(&contentType)

	fileItem := models.NewDriveItem()
	fileItem.SetName(&name)
	fileItem.SetAdditionalData(map[string]interface{}{
		"@microsoft.graph.conflictBehavior": "replace",
	})
	fileItem.SetFile(newFile)

	// When trying to upload to a personal OneDrive account, if the content is set the a error is returned -
	// Error creating file item: [child] A stream property 'content' has a value in the payload. In OData, stream property must not have a value, it must only use property annotations.
	// error: invalidRequest
	// error: [child] A stream property 'content' has a value in the payload. In OData, stream property must not have a value, it must only use property annotations.
	//
	// However, if the content is not set then a different error is returned -
	// Error creating file item: Cannot create a file without content
	// error: invalidRequest
	// error: Cannot create a file without content
	// fileItem.SetContent(data)

	createdFileItem, err := client.Drives().ByDriveId(driveID).Items().
		ByDriveItemId(*createdDriveItem.GetId()).Children().
		Post(context.Background(), fileItem, nil)
	if err != nil {
		fmt.Printf("Error creating file item: %v\n", err)
		printOdataError(err)
		return
	}

	// With Microsoft 365 OneDrive accounts I had to create the file in two steps to get it working,
	// first create the file item then upload the content.
	createdFileItem, err = client.Drives().ByDriveId(driveID).
		Items().ByDriveItemId(*createdFileItem.GetId()).
		Content().Put(context.Background(), data, nil)
	if err != nil {
		fmt.Printf("Error uploading file content: %v\n", err)
		printOdataError(err)
		return
	}

	fmt.Printf("Created file item: %v\n", *createdFileItem.GetId())
}

func createFolder(client *msgraphsdk.GraphServiceClient, driveID, parentID, name string) (models.DriveItemable, error) {
	folder := models.NewFolder()
	driveItem := models.NewDriveItem()
	driveItem.SetFolder(folder)
	driveItem.SetName(&name)
	driveItem.SetAdditionalData(map[string]interface{}{
		"@microsoft.graph.conflictBehavior": "replace",
	})

	createdDriveItem, err := client.Drives().ByDriveId(driveID).Items().
		ByDriveItemId(parentID).Children().
		Post(context.Background(), driveItem, nil)
	if err != nil {
		return nil, err
	}

	return createdDriveItem, nil
}

func printOdataError(err error) {
	var odataErr *odataerrors.ODataError
	switch {
	case errors.As(err, &odataErr):
		mainErr := odataErr.GetErrorEscaped()
		if code := mainErr.GetCode(); code != nil {
			fmt.Printf("error: %s\n", *code)
		}
		if message := mainErr.GetMessage(); message != nil {
			fmt.Printf("error: %s\n", *message)
		}
	default:
		fmt.Printf("%T > error: %#v", err, err)
	}
}
