package com.outlook.dev.service;

import org.json.simple.JSONObject;

import retrofit2.Call;
import retrofit2.http.Body;
import retrofit2.http.GET;
import retrofit2.http.POST;
import retrofit2.http.Path;
import retrofit2.http.Query;

public interface OutlookService {
	
	@GET("/v1.0/me")
	Call<OutlookUser> getCurrentUser();

	@GET("/v1.0/me/mailfolders/{folderid}/messages")
	Call<PagedResult<Message>> getMessages(
	  @Path("folderid") String folderId,
	  @Query("$orderby") String orderBy,
	  @Query("$select") String select,
	  @Query("$top") Integer maxResults
	);
	
	@GET("/v1.0/me/events")
	Call<PagedResult<Event>> getEvents(
	  @Query("$orderby") String orderBy,
	  @Query("$select") String select,
	  @Query("$top") Integer maxResults
	);
	
	@GET("/v1.0/me/contacts")
	Call<PagedResult<Contact>> getContacts(
		@Query("$orderby") String orderBy,
	  @Query("$select") String select,
	  @Query("$top") Integer maxResults
	);
	
	@POST("/api/v2.0/me/taskfolders({folder_id})/tasks")
	Call<POST> createTask(
		@Body JSONObject obj
	);
}
