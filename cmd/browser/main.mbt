///|
fn app() -> @rabbita.Val[@rabbita.Html] {
  let (model, emit) = @rabbita.create_state(initial_model(), update~)
  model.view(model => view(model, emit))
}

///|
fn main {
  @rabbita.new(app).mount("app")
}
