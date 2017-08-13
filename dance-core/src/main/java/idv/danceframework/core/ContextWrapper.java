package idv.danceframework.core;

public class ContextWrapper {
	private String title;
	private String field;

	public ContextWrapper(String title, String field) {
		this.title = title;
		this.field = field;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getField() {
		return field;
	}

	public void setField(String field) {
		this.field = field;
	}
}
